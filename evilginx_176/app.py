#!/usr/bin/env python3
from flask import Flask, request, g, render_template, jsonify, send_file
import sqlite3
import requests
import os
import json
from datetime import datetime, UTC
import tempfile
from openpyxl import Workbook
import subprocess

# Config
REMOTE_URL = os.environ.get('REMOTE_URL', 'http://172.20.14.176:8080/get_json')
DB_PATH = os.environ.get('CACHE_DB', 'cache.db')
FETCH_TIMEOUT = int(os.environ.get('FETCH_TIMEOUT', '20'))
PORT = int(os.environ.get('PORT', '5001'))
HOST = os.environ.get('HOST', '0.0.0.0')

app = Flask(__name__)

# ---------- Database helpers ----------
def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DB_PATH)
        db.row_factory = sqlite3.Row
    return db

def init_db():
    db = sqlite3.connect(DB_PATH)
    cur = db.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY,
            username TEXT,
            password TEXT,
            landing_url TEXT,
            tokens TEXT,
            fetched_at TEXT
        )
    ''')
    db.commit()
    db.close()

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

# ---------- NEW: FULL SYNC (no cache logic) ----------
def full_sync_with_remote():
    """Baixa dados remotos, atualiza, insere e remove usuários que não existem mais."""
    try:
        resp = requests.get(REMOTE_URL, timeout=FETCH_TIMEOUT)
    except Exception as e:
        return False, f'fetch-error: {e}'

    if resp.status_code != 200:
        return False, f'status-{resp.status_code}'

    try:
        remote_list = resp.json()
    except Exception as e:
        return False, f'json-error: {e}'

    if not isinstance(remote_list, list):
        if isinstance(remote_list, dict):
            remote_list = [remote_list]
        else:
            return False, 'unexpected-json-structure'

    # Normalizar IDs
    remote_ids = set()
    remote_map = {}

    for u in remote_list:
        if "id" not in u:
            continue
        uid = u["id"]
        remote_ids.add(uid)
        remote_map[uid] = u

    db = sqlite3.connect(DB_PATH)
    cur = db.cursor()

    # Buscar IDs locais
    cur.execute("SELECT id FROM users")
    local_ids = {row[0] for row in cur.fetchall()}

    # Determinar diferenças
    ids_to_delete = local_ids - remote_ids
    ids_to_insert_or_update = remote_ids

    # Remover usuários que não existem mais
    for uid in ids_to_delete:
        cur.execute("DELETE FROM users WHERE id = ?", (uid,))

    # Inserir/atualizar usuários recebidos
    for uid in ids_to_insert_or_update:
        u = remote_map[uid]
        cur.execute(
            '''
            INSERT INTO users (id, username, password, landing_url, tokens, fetched_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                username=excluded.username,
                password=excluded.password,
                landing_url=excluded.landing_url,
                tokens=excluded.tokens,
                fetched_at=excluded.fetched_at
            ''',
            (
                u.get("id"),
                u.get("username"),
                u.get("password"),
                u.get("landing_url"),
                json.dumps(u.get("tokens", {}), ensure_ascii=False),
                datetime.now(UTC).isoformat()
            )
        )

    db.commit()
    db.close()

    return True, f"sync-ok: remote={len(remote_ids)} local={len(local_ids)} deleted={len(ids_to_delete)}"


# ---------- Routes ----------

@app.route("/delete/cache", methods=["GET"])
def delete_cache():
    try:
        subprocess.run(["rm", "-f", "cache.db"], check=True)
        return jsonify({"status": "success", "message": "cache.db removido"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    ok, msg = full_sync_with_remote()
    return jsonify({'ok': ok, 'message': msg}), (200 if ok else 500)


@app.route('/api/users', methods=['GET'])
def api_users():
    db = get_db()
    cur = db.cursor()
    q_id = request.args.get('id')
    q_username = request.args.get('username')
    limit = int(request.args.get('limit') or 100)
    offset = int(request.args.get('offset') or 0)

    if q_id:
        try:
            q_id = int(q_id)
        except ValueError:
            return jsonify({'error': 'ID inválido'}), 400

        cur.execute('SELECT * FROM users WHERE id = ? LIMIT 1', (q_id,))
        row = cur.fetchone()
        if not row:
            return jsonify([])
        rowd = dict(row)
        rowd['tokens'] = json.loads(rowd['tokens'] or '{}')
        return jsonify([rowd])

    if q_username:
        cur.execute('SELECT * FROM users WHERE username LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?', (f'%{q_username}%', limit, offset))
    else:
        cur.execute('SELECT * FROM users ORDER BY id DESC LIMIT ? OFFSET ?', (limit, offset))

    rows = cur.fetchall()
    out = []
    for r in rows:
        rd = dict(r)
        rd['tokens'] = json.loads(rd['tokens'] or '{}')
        out.append(rd)
    return jsonify(out)


# ---------- Export XLSX ----------
@app.route('/api/export_xlsx', methods=['GET'])
def export_xlsx():
    db = get_db()
    cur = db.cursor()

    q_id = request.args.get('id')
    q_username = request.args.get('username')
    limit = int(request.args.get('limit') or 1000)
    offset = int(request.args.get('offset') or 0)

    if q_id:
        try:
            q_id = int(q_id)
        except ValueError:
            return jsonify({'error': 'ID inválido'}), 400
        cur.execute('SELECT id, username, password, landing_url FROM users WHERE id = ? LIMIT 1', (q_id,))
    elif q_username:
        cur.execute('SELECT id, username, password, landing_url FROM users WHERE username LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?', (f'%{q_username}%', limit, offset))
    else:
        cur.execute('SELECT id, username, password, landing_url FROM users ORDER BY id DESC LIMIT ? OFFSET ?', (limit, offset))

    rows = cur.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Users"

    ws.append(["ID", "Username", "Password", "Landing URL"])

    for r in rows:
        ws.append([r["id"], r["username"], r["password"], r["landing_url"]])

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.flush()
    tmp.seek(0)

    return send_file(tmp.name, as_attachment=True, download_name="users_export.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/')
def index():
    return render_template('index.html', remote=REMOTE_URL)


# ---------- Run ----------
if __name__ == '__main__':
    init_db()
    print(f"Starting app. Remote: {REMOTE_URL}. DB: {DB_PATH}. No cache mode.")
    app.run(host=HOST, port=PORT, debug=False)
