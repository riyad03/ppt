"""Start the web interface on the first free port and say where it is.

    python run_web.py

Running uvicorn directly fails with WinError 10048 when anything already holds
the port — most often an earlier run of this same server still alive in another
window — and the message says nothing about which address to try instead. This
picks a port that is actually free, prints the URL once, and gets out of the way.
"""

from __future__ import annotations

import socket
import sys

FIRST_PORT = 8000
ATTEMPTS = 20


def free_port(start: int, attempts: int) -> int:
    for port in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            # Without this, a port left in TIME_WAIT reports itself as free and
            # uvicorn then fails on the real bind.
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
            if probe.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise SystemExit(
        f"Aucun port libre entre {start} et {start + attempts - 1}. "
        "Fermez une fenêtre qui fait déjà tourner le serveur."
    )


def main() -> None:
    try:
        import uvicorn
    except ImportError:
        raise SystemExit(
            "FastAPI et uvicorn ne sont pas installés dans cet environnement.\n"
            "  pip install fastapi \"uvicorn[standard]\" python-multipart"
        )

    port = free_port(FIRST_PORT, ATTEMPTS)
    url = f"http://127.0.0.1:{port}"
    print("=" * 56)
    print(f"  Atlas Guardian — interface web")
    print(f"  {url}")
    if port != FIRST_PORT:
        print(f"  (le port {FIRST_PORT} était occupé)")
    print(f"  Ctrl+C pour arrêter. Gardez cette fenêtre ouverte :")
    print(f"  le suivi des générations vit dans ce processus.")
    print("=" * 56)
    sys.stdout.flush()

    uvicorn.run("web.server:app", host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    main()
