#!/usr/bin/env python3
"""Local GET/POST echo API for the templates; binds to loopback only."""
import argparse
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qsl, urlsplit


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def do_GET(self):
        self.respond()

    def do_POST(self):
        self.respond()

    def respond(self):
        size = int(self.headers.get('Content-Length', '0'))
        if size > 10 * 1024 * 1024:
            self.send_error(413)
            return
        body = self.rfile.read(size).decode('utf-8', errors='replace')
        parts = urlsplit(self.path)
        payload = json.dumps({'method': self.command, 'path': parts.path,
                              'query': parse_qsl(parts.query, keep_blank_values=True),
                              'body': body}, ensure_ascii=False).encode('utf-8')
        self.send_response(200 if parts.path in ('/health', '/echo') else 404)
        self.send_header('Content-Type', 'application/json; charset=UTF-8')
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format_, *args):
        # Do not print potentially sensitive query strings or header/body content.
        pass


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8080)
    args = parser.parse_args()
    with ThreadingHTTPServer(('127.0.0.1', args.port), Handler) as server:
        print('Local test API: http://127.0.0.1:%d (Ctrl+C to stop)' % server.server_port, flush=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    main()
