#!/usr/bin/env python3
"""Simple HTTPS server for Office Add-in development."""
import http.server
import ssl
import os

PORT = 3000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CERT_FILE = os.path.join(BASE_DIR, 'certs', 'server.crt')
KEY_FILE = os.path.join(BASE_DIR, 'certs', 'server.key')

os.chdir(BASE_DIR)

handler = http.server.SimpleHTTPRequestHandler
handler.extensions_map.update({
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.html': 'text/html',
    '.json': 'application/json',
    '.png': 'image/png',
    '.xml': 'application/xml',
})

httpd = http.server.HTTPServer(('localhost', PORT), handler)

ssl_ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
ssl_ctx.load_cert_chain(CERT_FILE, KEY_FILE)
httpd.socket = ssl_ctx.wrap_socket(httpd.socket, server_side=True)

print(f'Serving on https://localhost:{PORT}')
print('Press Ctrl+C to stop')
httpd.serve_forever()
