#!/usr/bin/env python3
from flask import Flask, request, g, render_template, jsonify, send_file
import sqlite3
import requests
import os
import json
from datetime import datetime, UTC
import tempfile
from openpyxl import Workbook

# Config
REMOTE_URL = os.environ.get('REMOTE_URL', 'http://172.20.14.153:8080/get_json')
DB_PATH = os.environ.get('CACHE_DB', 'cache.db')
FETCH_TIMEOUT = int(os.environ.get('FETCH_TIMEOUT', '20'))
PORT = int(os.environ.get('PORT', '5004'))
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

# ---------- Sync logic (no cache) ----------

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
            json.dumps(user.get('tokens', []), ensure_ascii=False),
            datetime.now(UTC).isoformat()
        )
    )

def fetch_and_cache():
    """Fetch remoto e sincroniza o DB completamente (sem cache TTL)."""
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

    api_ids = {u["id"] for u in payload if "id" in u}

    db = sqlite3.connect(DB_PATH)
    try:
        cur = db.cursor()
        cur.execute("SELECT id FROM users")
        db_ids = {row[0] for row in cur.fetchall()}

        to_delete = db_ids - api_ids
        if to_delete:
            cur.executemany("DELETE FROM users WHERE id = ?", [(i,) for i in to_delete])

        for u in payload:
            if 'id' not in u:
                continue
            upsert_user(db, u)

        db.commit()
    except Exception as e:
        return False, f'db-error: {e}'
    finally:
        db.close()

    return True, f'synchronized {len(payload)} items'

# ---------- Routes ----------

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
        rowd['tokens'] = json.loads(rowd['tokens'] or '[]')
        return jsonify([rowd])

    if q_username:
        cur.execute('SELECT * FROM users WHERE username LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?', (f'%{q_username}%', limit, offset))
    else:
        cur.execute('SELECT * FROM users ORDER BY id DESC LIMIT ? OFFSET ?', (limit, offset))

    rows = cur.fetchall()
    out = []
    for r in rows:
        rd = dict(r)
        rd['tokens'] = json.loads(rd['tokens'] or '[]')
        out.append(rd)
    return jsonify(out)

# Export XLSX
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
