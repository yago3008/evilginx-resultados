#!/usr/bin/env python3
from flask import Flask, request, g, render_template, jsonify, send_file
import sqlite3
import requests
import os
import json
from datetime import datetime, timedelta, UTC
import threading
import tempfile
from openpyxl import Workbook
import subprocess

# Config
REMOTE_URL = os.environ.get('REMOTE_URL', 'http://172.20.14.176:8080/get_json')
DB_PATH = os.environ.get('CACHE_DB', 'cache.db')
CACHE_TTL_SECONDS = int(os.environ.get('CACHE_TTL_SECONDS', '10'))
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

# ---------- Fetch & cache logic ----------
def upsert_user(db_conn, user):
    cur = db_conn.cursor()
    cur.execute(
        '''INSERT INTO users (id, username, password, landing_url, tokens, fetched_at)
           VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(id) DO UPDATE SET
             username=excluded.username,
             password=excluded.password,
             landing_url=excluded.landing_url,
             tokens=excluded.tokens,
             fetched_at=excluded.fetched_at
        ''',
        (
            user.get('id'),
            user.get('username'),
            user.get('password'),
            user.get('landing_url'),
            json.dumps(user.get('tokens', {}), ensure_ascii=False),
            datetime.now(UTC).isoformat()  # ✅ corrigido
        )
    )

def fetch_and_cache():
    """Fetch remote JSON and upsert into sqlite. Returns (ok, message)."""
    try:
        resp = requests.get(REMOTE_URL, timeout=FETCH_TIMEOUT)
    except Exception as e:
        return False, f'fetch-error: {e}'
    if resp.status_code != 200:
        return False, f'status-{resp.status_code}'

    try:
        payload = resp.json()
    except Exception as e:
        return False, f'json-error: {e}'

    if not isinstance(payload, list):
        if isinstance(payload, dict):
            payload = [payload]
        else:
            return False, 'unexpected-json-structure'

    db = sqlite3.connect(DB_PATH)
    try:
        seen = set()
        for u in payload:
            if 'id' not in u:
                continue
            key = u.get('username', u.get('password'))
            if key in seen:
                continue
            seen.add(key)
            upsert_user(db, u)
        db.commit()
    except Exception as e:
        return False, f'db-error: {e}'
    finally:
        db.close()

    return True, f'cached {len(payload)} items'

def last_fetch_time():
    db = get_db()
    cur = db.cursor()
    cur.execute('SELECT fetched_at FROM users ORDER BY fetched_at DESC LIMIT 1')
    row = cur.fetchone()
    if not row or not row['fetched_at']:
        return None
    return datetime.fromisoformat(row['fetched_at'])

def cache_is_stale():
    t = last_fetch_time()
    if t is None:
        return True
    return (datetime.now(UTC) - t) > timedelta(seconds=CACHE_TTL_SECONDS)

# ---------- Routes ----------

@app.route("/delete/cache", methods=["GET"])
def delete_cache():
    try:
        # Executa o comando rm usando subprocess de forma segura
        subprocess.run(["rm", "-f", "cache.db"], check=True)

        return jsonify({"status": "success", "message": "cache.db removido"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500



@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    ok, msg = fetch_and_cache()
    status = 200 if ok else 500
    return jsonify({'ok': ok, 'message': msg}), status


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


# ✅ NOVA ROTA: exportar XLSX
@app.route('/api/export_xlsx', methods=['GET'])
def export_xlsx():
    """Gera e retorna um arquivo XLSX com os dados atuais."""
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

    # Cria o arquivo Excel temporário
    wb = Workbook()
    ws = wb.active
    ws.title = "Users"

    # Cabeçalhos
    ws.append(["ID", "Username", "Password", "Landing URL"])

    # Dados
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
    print(f'Starting app. Remote: {REMOTE_URL}. DB: {DB_PATH}. Cache TTL {CACHE_TTL_SECONDS}s')
    app.run(host=HOST, port=PORT, debug=False)
