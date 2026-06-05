#!/usr/bin/env python3
"""
HMV Ingenieros — Sistema de Viáticos
Servidor web para despliegue en Render.com
"""
import http.server, json, os, subprocess, sys, traceback
from http.server import HTTPServer, BaseHTTPRequestHandler

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILL_PY    = os.path.join(SCRIPT_DIR, 'fill_template.py')
HTML_FILE  = os.path.join(SCRIPT_DIR, 'viaticos-hmv.html')

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[HMV] {self.address_string()} — {format % args}", flush=True)

    def send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_HEAD(self):
        """Render health check uses HEAD — respond 200 OK."""
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def do_GET(self):
        path = self.path.split('?')[0]

        # Health check endpoint
        if path == '/health':
            self._json(200, {'status': 'ok', 'python': sys.version})
            return

        # Diagnóstico: verifica soffice y fill_template
        if path == '/diagnostico':
            resultado = {}
            # Verificar soffice
            try:
                r = subprocess.run(['soffice', '--version'],
                                   capture_output=True, text=True, timeout=10)
                resultado['soffice'] = r.stdout.strip() or r.stderr.strip()
            except Exception as e:
                resultado['soffice'] = f'ERROR: {e}'
            # Verificar fill_template.py
            resultado['fill_template'] = os.path.exists(FILL_PY)
            resultado['plantilla']     = os.path.exists(
                os.path.join(SCRIPT_DIR, 'plantilla.docx'))
            resultado['python_docx']   = self._check_import('docx')
            resultado['lxml']          = self._check_import('lxml')
            self._json(200, resultado)
            return

        # Servir HTML principal
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
                self._json(500, {'error': str(e)})
            return

        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        if self.path == '/generar-pdf':
            length  = int(self.headers.get('Content-Length', 0))
            payload = self.rfile.read(length)
            try:
                data = json.loads(payload)
                print(f"[HMV] Generando PDF para: {data.get('nombre','?')}", flush=True)

                result = subprocess.run(
                    [sys.executable, FILL_PY, json.dumps(data)],
                    capture_output=True, text=True, timeout=120,
                    cwd=SCRIPT_DIR
                )

                if result.returncode != 0:
                    raise Exception(f"fill_template error: {result.stderr or result.stdout}")

                b64pdf = result.stdout.strip()
                if not b64pdf:
                    raise Exception("fill_template no generó salida")

                self._json(200, {'ok': True, 'pdf': b64pdf})

            except subprocess.TimeoutExpired:
                self._json(500, {'ok': False, 'error': 'Tiempo de espera agotado generando el PDF'})
            except Exception as e:
                print(f"[HMV] ERROR: {traceback.format_exc()}", flush=True)
                self._json(500, {'ok': False, 'error': str(e)})
            return

        self.send_response(404)
        self.end_headers()

    def _json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.send_cors()
        self.end_headers()
        self.wfile.write(body)

    def _check_import(self, mod):
        try:
            __import__(mod)
            return True
        except ImportError:
            return False

if __name__ == '__main__':
    port   = int(os.environ.get('PORT', 10000))
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f"\n{'='*50}", flush=True)
    print(f"  HMV Sistema de Viáticos — Puerto {port}", flush=True)
    print(f"  fill_template: {FILL_PY}", flush=True)
    print(f"  HTML: {HTML_FILE}", flush=True)
    print(f"{'='*50}\n", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
