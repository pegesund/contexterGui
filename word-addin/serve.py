#!/usr/bin/env python3
"""HTTPS server for Word Add-in. Proxies /context and /reply to Rust app."""
import http.server, ssl, os, urllib.request, json

RUST_URL = "http://localhost:52525"

os.chdir(os.path.dirname(os.path.abspath(__file__)))

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path in ("/context", "/changed", "/deleted", "/reset"):
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                req = urllib.request.Request(RUST_URL + self.path, data=body,
                    headers={"Content-Type": "application/json"}, method="POST")
                resp = urllib.request.urlopen(req, timeout=2)
                result = resp.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", len(result))
                self.end_headers()
                self.wfile.write(result)
            except Exception as e:
                msg = json.dumps({"error": str(e)}).encode()
                self.send_response(502)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", len(msg))
                self.end_headers()
                self.wfile.write(msg)
        else:
            self.send_response(404)
            self.end_headers()

    def do_GET(self):
        if self.path == "/reply":
            try:
                resp = urllib.request.urlopen(RUST_URL + "/reply", timeout=2)
                result = resp.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", len(result))
                self.end_headers()
                self.wfile.write(result)
            except Exception as e:
                msg = b"{}"
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", len(msg))
                self.end_headers()
                self.wfile.write(msg)
        elif self.path == "/ping":
            try:
                resp = urllib.request.urlopen(RUST_URL + "/ping", timeout=2)
                result = resp.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/plain")
                self.send_header("Content-Length", len(result))
                self.end_headers()
                self.wfile.write(result)
            except:
                self.send_response(502)
                self.end_headers()
        else:
            super().do_GET()

    def log_message(self, format, *args):
        # Suppress request logs except errors
        if args and "404" not in str(args[0]) and "502" not in str(args[0]):
            return
        super().log_message(format, *args)

server = http.server.HTTPServer(("localhost", 3000), Handler)
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
ctx.load_cert_chain("cert.pem", "key.pem")
server.socket = ctx.wrap_socket(server.socket, server_side=True)

print("Word Add-in server at https://localhost:3000")
print("Proxying /context and /reply to Rust app at", RUST_URL)
server.serve_forever()
