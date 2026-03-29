from __future__ import annotations

import argparse
import http.server
import socketserver
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from refresh_dashboard_data import refresh_dashboard_data  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Serve the GitHub dashboard bundle locally.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--workbook", help="Optional local workbook path. If omitted, the server searches the parent folder for .xlsm/.xlsx files.")
    return parser.parse_args()


def find_default_workbook(bundle_dir: Path) -> Path | None:
    parent = bundle_dir.parent
    for pattern in ("*.xlsm", "*.xlsx"):
        matches = sorted(parent.glob(pattern))
        if matches:
            return matches[0]
    return None


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    workbook_path: Path | None = None
    bundle_dir: Path = BUNDLE_DIR

    def _refresh_if_needed(self, force: bool = False) -> None:
        if self.workbook_path is None:
            return
        output_path = self.bundle_dir / "dashboard_data.json"
        if force or (not output_path.exists()) or self.workbook_path.stat().st_mtime > output_path.stat().st_mtime:
            refresh_dashboard_data(workbook=str(self.workbook_path), output=output_path)

    def do_GET(self) -> None:  # noqa: N802
        if self.path.startswith("/api/dashboard-data"):
            self._refresh_if_needed(force=True)
            body = (self.bundle_dir / "dashboard_data.json").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return

        if self.path.startswith("/dashboard_data.json"):
            self._refresh_if_needed(force=True)

        super().do_GET()

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve() if args.workbook else find_default_workbook(BUNDLE_DIR)
    if workbook_path and not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    DashboardHandler.workbook_path = workbook_path
    DashboardHandler.bundle_dir = BUNDLE_DIR

    if workbook_path:
        refresh_dashboard_data(workbook=str(workbook_path), output=BUNDLE_DIR / "dashboard_data.json")

    handler = lambda *a, **kw: DashboardHandler(*a, directory=str(BUNDLE_DIR), **kw)  # noqa: E731
    with socketserver.TCPServer((args.host, args.port), handler) as httpd:
        print(f"Serving dashboard on http://{args.host}:{args.port}")
        if workbook_path:
            print(f"Watching workbook: {workbook_path}")
        httpd.serve_forever()


if __name__ == "__main__":
    main()
