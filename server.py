"""Contact book: Python 3.10+, SQLite, real XLSX import/export."""
import io
import json
import os
import sqlite3
import traceback
from contextlib import closing
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent
DB = Path(os.environ.get('DATABASE_PATH', str(ROOT / 'data' / 'contacts.db')))
DB.parent.mkdir(parents=True, exist_ok=True)
TOKEN = os.environ.get('APP_PASSWORD', '')
TYPES = ['phone', 'email', 'social', 'address']

def connect():
    db = sqlite3.connect(DB, timeout=20)
    db.execute('CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY, data TEXT NOT NULL)')
    return db

def validate(item):
    if not isinstance(item, dict):
        raise ValueError('联系人格式错误')
    name = str(item.get('name', '')).strip()
    if not name or len(name) > 100:
        raise ValueError('姓名必填，最多100字')
    methods = item.get('methods', [])
    if not isinstance(methods, list) or len(methods) > 100:
        raise ValueError('联系方式格式错误或超过100条')
    clean = []
    for m in methods:
        if not isinstance(m, dict) or m.get('type') not in TYPES:
            raise ValueError('联系方式类型错误')
        value = str(m.get('value', '')).strip()
        if not value or len(value) > 1000:
            raise ValueError('联系方式不能为空，最多1000字')
        clean.append({'type': m['type'], 'value': value})
        if m['type'] == 'social' and m.get('platform'):
            platform = str(m['platform']).strip()
            if len(platform) > 50: raise ValueError('社交软件名称最多50字')
            if platform: clean[-1]['platform'] = platform
    notes = str(item.get('notes', '')).strip()
    if len(notes) > 5000:
        raise ValueError('备注最多5000字')
    return dict(name=name, notes=notes, favorite=bool(item.get('favorite', False)), methods=clean)

def all_contacts(db):
    return [dict(json.loads(data), id=i) for i, data in db.execute('SELECT id,data FROM contacts ORDER BY id DESC')]

def export_xlsx(items):
    wb = Workbook(); ws = wb.active; ws.title = 'Contacts'
    counts = {t: max([sum(m['type'] == t for m in c['methods']) for c in items] + [1]) for t in TYPES}
    columns = [(t, i) for t in TYPES for i in range(1, counts[t]+1)]
    ws.append(['name', 'favorite', 'notes'] + [f'{t}_{i}' for t, i in columns])
    for c in items:
        values = {t: [(m.get('platform', '') + '：' + m['value']) if t == 'social' and m.get('platform') else m['value'] for m in c['methods'] if m['type'] == t] for t in TYPES}
        row = [c['name'], int(c['favorite']), c['notes']] + [values[t][i-1] if i <= len(values[t]) else '' for t, i in columns]
        ws.append(row)
        # Force user strings to text: no Excel formula execution.
        for cell in ws[ws.max_row]:
            if isinstance(cell.value, str): cell.data_type = 's'
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    for column in ws.columns: ws.column_dimensions[column[0].column_letter].width = 25
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def import_xlsx(data):
    wb = load_workbook(io.BytesIO(data), read_only=False, data_only=False)
    ws = wb.active
    # Streaming-generated XLSX may omit worksheet dimensions entirely.
    if (ws.max_row or 0) > 10001 or (ws.max_column or 0) > 403:
        raise ValueError('最多10000位联系人、403列')
    rows = iter(ws.values); headers = list(next(rows, []))
    if len(headers) > 403:
        raise ValueError('最多403列')
    if not headers or headers[:3] != ['name', 'favorite', 'notes']:
        raise ValueError('请使用导出的Excel作为模板，前三列为name、favorite、notes')
    if len(set(headers)) != len(headers): raise ValueError('列名不能重复')
    for h in headers[3:]:
        parts = str(h).rsplit('_', 1)
        if len(parts) != 2 or parts[0] not in TYPES + ['socialplatform'] or not parts[1].isdigit() or int(parts[1]) < 1:
            raise ValueError('未知联系方式列：' + str(h))
    items = []
    for index, row in enumerate(rows, 2):
        if index > 10001 or len(row) > 403:
            raise ValueError('最多10000位联系人、403列')
        if all(v is None for v in row): continue
        try:
            fav = row[1]
            if str(fav or '').lower() not in ('', '0', '1', 'true', 'false'): raise ValueError('favorite必须为0或1')
            item = dict(name=row[0] or '', favorite=str(fav).lower() in ('1', 'true'), notes=row[2] or '', methods=[])
            for h, value in zip(headers[3:], row[3:]):
                kind, number = h.rsplit('_', 1)
                if kind == 'socialplatform': continue
                if value is not None and str(value).strip():
                    method = {'type': kind, 'value': str(value)}
                    platform_col = 'socialplatform_' + number
                    if kind == 'social' and platform_col in headers:
                        platform = row[headers.index(platform_col)]
                        if platform: method['platform'] = str(platform)
                    elif kind == 'social' and '：' in str(value):
                        platform, account = str(value).split('：', 1)
                        if platform.strip() and account.strip():
                            method['platform'] = platform.strip()
                            method['value'] = account.strip()
                    item['methods'].append(method)
            items.append(validate(item))
        except ValueError as e: raise ValueError(f'第{index}行：{e}') from e
    wb.close(); return items

class Handler(BaseHTTPRequestHandler):
    def reply(self, code, data, content_type='application/json; charset=utf-8'):
        if not isinstance(data, bytes): data = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code); self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data))); self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Cache-Control', 'no-store'); self.end_headers(); self.wfile.write(data)

    def do_GET(self): self.handle_request()
    def do_POST(self): self.handle_request()
    def do_PUT(self): self.handle_request()
    def do_DELETE(self): self.handle_request()

    def handle_request(self):
        path = urlparse(self.path).path
        if not path.startswith('/api/'):
            files = {'/': ('index.html', 'text/html; charset=utf-8'), '/app.js': ('app.js', 'text/javascript; charset=utf-8'), '/style.css': ('style.css', 'text/css; charset=utf-8')}
            if path not in files or self.command != 'GET': return self.reply(404, {'error': 'Not found'})
            filename, mime = files[path]; return self.reply(200, (ROOT/'static'/filename).read_bytes(), mime)
        if TOKEN and self.headers.get('Authorization') != 'Bearer ' + TOKEN:
            return self.reply(401, {'error': '访问密码不正确'})
        try:
            size = int(self.headers.get('Content-Length', 0))
            if size > 10*1024*1024: return self.reply(413, {'error': '文件不能超过10MB'})
            body = self.rfile.read(size)
            with closing(connect()) as db:
                if path == '/api/contacts' and self.command == 'GET': return self.reply(200, all_contacts(db))
                if path == '/api/export' and self.command == 'GET':
                    return self.reply(200, export_xlsx(all_contacts(db)), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                if path == '/api/import' and self.command == 'POST':
                    items = import_xlsx(body)
                    db.executemany('INSERT INTO contacts(data) VALUES(?)', [(json.dumps(c, ensure_ascii=False),) for c in items])
                    db.commit(); return self.reply(200, {'imported': len(items)})
                if path == '/api/contacts' and self.command == 'POST':
                    item = validate(json.loads(body)); cur = db.execute('INSERT INTO contacts(data) VALUES(?)', (json.dumps(item, ensure_ascii=False),))
                    db.commit(); return self.reply(201, dict(item, id=cur.lastrowid))
                if path.startswith('/api/contacts/'):
                    ident = int(path.rsplit('/', 1)[1])
                    if not db.execute('SELECT 1 FROM contacts WHERE id=?', (ident,)).fetchone(): return self.reply(404, {'error': '联系人不存在'})
                    if self.command == 'DELETE': db.execute('DELETE FROM contacts WHERE id=?', (ident,))
                    elif self.command == 'PUT':
                        item = validate(json.loads(body)); db.execute('UPDATE contacts SET data=? WHERE id=?', (json.dumps(item, ensure_ascii=False), ident))
                    else: return self.reply(405, {'error': 'Method not allowed'})
                    db.commit(); return self.reply(200, {'ok': True})
                return self.reply(404, {'error': 'Not found'})
        except (ValueError, KeyError, TypeError) as e:
            traceback.print_exc()
            self.reply(400, {'error': str(e)})
        except Exception:
            traceback.print_exc()
            self.reply(400, {'error': '处理失败，请检查文件格式或稍后重试'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8000'))
    print(f'Contact book v1.0.2: http://localhost:{port}', flush=True)
    print(f'Running file: {Path(__file__).resolve()}', flush=True)
    ThreadingHTTPServer((os.environ.get('HOST', '127.0.0.1'), port), Handler).serve_forever()
