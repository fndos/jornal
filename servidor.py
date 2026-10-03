#!/usr/bin/env python3
"""Servidor local para desarrollo.

Igual que `python -m http.server`, pero pidiendo al navegador que no guarde
nada en cache: si no, se queda con el app.js viejo y uno cree que los cambios
no se aplicaron.

    python3 servidor.py
"""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import os

PUERTO = 8000


class SinCache(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, formato, *args):
        if "200" not in formato % args:
            super().log_message(formato, *args)


if __name__ == "__main__":
    os.chdir(Path(__file__).parent)
    print(f"Jornal en http://127.0.0.1:{PUERTO}/  (sin cache)")
    ThreadingHTTPServer(("127.0.0.1", PUERTO), SinCache).serve_forever()
