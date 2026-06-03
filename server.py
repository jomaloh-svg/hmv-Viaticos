#!/usr/bin/env python3
"""
HMV Ingenieros — Sistema de Viáticos
Servidor web para despliegue en Render.com
"""
import http.server, json, os, subprocess, sys
from http.server import HTTPServer, BaseHTTPRequestHandler

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLA  = os.path.join(SCRIPT_DIR, 'plantilla.docx')
FILL_PY    = os.path.join(SCRIPT_DIR, 'fill_template.py')
HTML_FILE  = os.path.join(SCRIPT_DIR, 'viaticos-hmv.html')

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[HMV] {self.address_string()} — {format % args}", flush=True)

    def send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def do_GET(self):
        path = self.path.split('?')[0]
        if path in ('/', '/index.html'):
            try:
                with open(HTML_FILE, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', len(content))
                self.end_headers()
                self.wfile.write(content)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/generar-pdf':
            length  = int(self.headers.get('Content-Length', 0))
            payload = self.rfile.read(length)
            try:
                data   = json.loads(payload)
                b64pdf = subprocess.check_output(
                    [sys.executable, FILL_PY, json.dumps(data)],
                    cwd=SCRIPT_DIR
                ).decode('utf-8').strip()
                resp = json.dumps({'ok': True, 'pdf': b64pdf}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', len(resp))
                self.send_cors()
                self.end_headers()
                self.wfile.write(resp)
            except Exception as e:
                err = json.dumps({'ok': False, 'error': str(e)}).encode('utf-8')
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', len(err))
                self.send_cors()
                self.end_headers()
                self.wfile.write(err)
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    host   = '0.0.0.0'
    server = HTTPServer((host, port), Handler)
    print(f"\n{'='*50}", flush=True)
    print(f"  HMV Sistema de Viáticos — Puerto {port}", flush=True)
    print(f"{'='*50}\n", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
