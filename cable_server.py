"""
Cable Schedule Conversion Server
Start:   python cable_server.py
Open:    http://127.0.0.1:8765/
"""

import sys, os, re, json, io
from http.server import HTTPServer, SimpleHTTPRequestHandler
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    os.system('pip install pdfplumber -q')
    import pdfplumber

PORT = 8765
DIR = os.path.dirname(os.path.abspath(__file__))

def is_valid_cable_number(val):
    if not val: return False
    s = str(val).strip()
    if not s: return False
    if s.lower() in ('cable number', 'cable no', 'cable id'): return False
    if re.match(r'^[A-Z]{1,6}-\d+', s, re.I): return True
    return bool(re.match(r'^[A-Z]{2,6}\s*\d+', s, re.I))

def is_section_header(val):
    if not val: return False
    s = str(val).strip().upper()
    return s in ('AC POWER', 'DC POWER', 'COAXIAL', 'FIBER OPTIC', 'FIBRE OPTIC',
                 'ANTENNA / COAXIAL NETWORK', 'WIRELESS NETWORK')

def has_cable_label_separator(val):
    if not val: return False
    return ' - ' in re.sub(r'\s+', ' ', str(val).strip())

def normalize_label(val):
    if not val: return ''
    return re.sub(r'\s+', ' ', str(val).strip())

def find_cable_headers(row):
    cn_col = cl_col = -1
    for ci, cell in enumerate(row):
        if cell is None: continue
        hl = str(cell).strip().lower()
        if hl in ('cable number', 'cable no.', 'cable no', 'cable id', 'cable #'):
            cn_col = ci
        if hl == 'cable label' or hl.startswith('cable label') or ('cable label' in hl):
            cl_col = ci
    return cn_col, cl_col

def extract_from_pdf_bytes(pdf_bytes):
    results = []
    errors = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            if not tables:
                errors.append(f"Page {page_num}: no tables found")
                continue
            for table in tables:
                if not table or len(table) < 2: continue
                header_row = None
                cn_col = cl_col = -1
                for ri, row in enumerate(table):
                    if not row: continue
                    c, l = find_cable_headers(row)
                    if c >= 0:
                        header_row = ri
                        cn_col = c
                        cl_col = l if l >= 0 else cl_col
                        break
                if header_row is None: continue
                last_cn = ''
                for ri in range(header_row + 1, len(table)):
                    row = table[ri]
                    if not row: continue
                    raw_cn = str(row[cn_col]).strip() if cn_col < len(row) and row[cn_col] is not None else ''
                    raw_cl = str(row[cl_col]).strip() if cl_col >= 0 and cl_col < len(row) and row[cl_col] is not None else ''
                    cn = raw_cn.strip()
                    cl = normalize_label(raw_cl)
                    if not cl: continue
                    if is_section_header(cl) or is_section_header(cn):
                        last_cn = ''
                        continue
                    if not cn:
                        if last_cn:
                            cn = last_cn
                        else:
                            continue
                    else:
                        if not is_valid_cable_number(cn):
                            last_cn = ''
                            continue
                        last_cn = cn
                    if not has_cable_label_separator(cl):
                        errors.append(f"Page {page_num}, row {ri+1}: '{cn}' label missing ' - ' separator: '{cl}'")
                    results.append({"cableNumber": cn, "cableLabel": cl})
    return results, errors

class CableServer(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIR, **kwargs)

    def _json_response(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())

    def do_POST(self):
        if self.path == '/convert':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                pdf_bytes = self.rfile.read(content_length)
                if not pdf_bytes:
                    self._json_response({"error": "No data received"}, 400)
                    return
                results, errors = extract_from_pdf_bytes(pdf_bytes)
                prefixes = {}
                for r in results:
                    m = re.match(r'^([A-Z]+)', r['cableNumber'], re.I)
                    if m:
                        k = m.group(1).upper()
                        prefixes[k] = prefixes.get(k, 0) + 1
                self._json_response({
                    "success": True, "total": len(results),
                    "issues": len(errors), "cables": results,
                    "errors": errors, "prefixes": prefixes
                })
            except Exception as e:
                self._json_response({"error": str(e)}, 500)
        else:
            self._json_response({"error": "Not found"}, 404)

    def do_GET(self):
        if self.path == '/ping':
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.end_headers()
            self.wfile.write(b'ok')
        else:
            super().do_GET()

    def log_message(self, format, *args):
        try:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {args}")
        except: pass

if __name__ == '__main__':
    server = HTTPServer(('127.0.0.1', PORT), CableServer)
    print(f"Open: http://127.0.0.1:{PORT}/")
    print(f"Press Ctrl+C to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.server_close()
