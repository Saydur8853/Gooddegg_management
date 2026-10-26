from __future__ import annotations

import threading

import webview

from app import HOST, PORT, InventoryHandler
from http.server import ThreadingHTTPServer


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), InventoryHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    webview.create_window(
        "Gooddegg Inventory Management",
        f"http://{HOST}:{PORT}",
        width=1280,
        height=850,
        min_size=(960, 650),
        confirm_close=True,
    )
    try:
        webview.start(gui="edgechromium", debug=False)
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
