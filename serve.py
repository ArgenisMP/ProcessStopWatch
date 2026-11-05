# Local test server: python serve.py  ->  http://localhost:8080
import http.server, socketserver
class H(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map, '.webmanifest': 'application/manifest+json', '.js': 'text/javascript'}
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(('', 8080), H) as s:
    print('Serving on http://localhost:8080'); s.serve_forever()
