"""
ClickUp Public-Share -> Excel exporter (web UI).

Run:  python app.py            -> open http://127.0.0.1:5000
Or run the packaged clickup-scraper.exe (browser opens automatically).
"""
from __future__ import annotations

import io
import os
import re
import socket
import sys
import threading
import time
import webbrowser
from datetime import datetime

from flask import (Flask, Response, jsonify, render_template, request,
                   send_file)

from exporter import build_workbook
from scraper import ScrapeError, scrape


def _resource_path(rel: str) -> str:
    """Resolve a bundled resource path (works both as script and frozen exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


app = Flask(__name__, template_folder=_resource_path("templates"))


def _safe_filename(name: str) -> str:
    name = re.sub(r"[^\w฀-๿\- ]+", "", name or "clickup").strip()
    name = re.sub(r"\s+", "_", name) or "clickup"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return f"{name}_{stamp}.xlsx"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/export", methods=["POST"])
def export():
    data = request.get_json(silent=True) or request.form
    url = (data.get("url") or "").strip()
    try:
        result = scrape(url)
    except ScrapeError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=f"เกิดข้อผิดพลาดไม่คาดคิด: {exc}"), 500

    xlsx = build_workbook(result)
    filename = _safe_filename(result.get("view_name"))
    resp = send_file(
        io.BytesIO(xlsx),
        mimetype=("application/vnd.openxmlformats-officedocument"
                  ".spreadsheetml.sheet"),
        as_attachment=True,
        download_name=filename,
    )
    resp.headers["X-Task-Count"] = str(len(result.get("tasks", [])))
    resp.headers["X-Error-Count"] = str(len(result.get("errors", [])))
    return resp


@app.route("/preview", methods=["POST"])
def preview():
    """Scrape and return a JSON summary (no download) for quick verification."""
    data = request.get_json(silent=True) or request.form
    url = (data.get("url") or "").strip()
    try:
        result = scrape(url)
    except ScrapeError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=f"เกิดข้อผิดพลาดไม่คาดคิด: {exc}"), 500

    tasks = result.get("tasks", [])
    total_comments = sum(len(t.get("comments", []) or []) for t in tasks)
    total_subtasks = sum(len(t.get("subtasks", []) or []) for t in tasks)
    return jsonify(
        ok=True,
        view_name=result.get("view_name"),
        task_count=len(tasks),
        comment_count=total_comments,
        subtask_count=total_subtasks,
        error_count=len(result.get("errors", [])),
        sample=[{"id": t.get("id"), "name": t.get("name"),
                 "comments": len(t.get("comments", []) or [])}
                for t in tasks[:5]],
    )


def _pick_port(preferred: int = 5000) -> int:
    """Use the preferred port if free, otherwise let the OS pick one."""
    for port in (preferred, 0):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                s.bind(("127.0.0.1", port))
                return s.getsockname()[1]
        except OSError:
            continue
    return preferred


def main():
    port = _pick_port(5000)
    url = f"http://127.0.0.1:{port}"
    # Open the browser shortly after the server starts.
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    print("=" * 56)
    print("  ClickUp -> Excel Exporter")
    print(f"  เปิดใช้งานที่: {url}")
    print("  ปิดโปรแกรม: กดปิดหน้าต่างนี้ หรือ Ctrl+C")
    print("=" * 56)
    app.run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
