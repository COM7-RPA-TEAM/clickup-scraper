"""
ClickUp public-share scraper.

Pulls every task in a public ClickUp share link (sharing.clickup.com) together
with all of its fields and comments by calling the same public API the share
page uses internally. No login / API token required.

Flow (reverse-engineered from the share page network traffic):
  1. GET  id.app.clickup.com/shard/v1/handshake/{workspace}      -> frontdoor host
  2. GET  {frontdoor}/view/v1/{ws}/public/view/{view}?token=...  -> all task ids
  3. GET  {frontdoor}/view/v1/{ws}/public/view/{view}/task/{id}?token=...
                                                                  -> full task + comments
"""
from __future__ import annotations

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import requests

HANDSHAKE_URL = "https://id.app.clickup.com/shard/v1/handshake/{ws}"

# Only header the public endpoints actually require.
BASE_HEADERS = {
    "x-csrf": "1",
    "accept": "application/json",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
}


class ScrapeError(Exception):
    """Raised when a share link can't be parsed or the API rejects us."""


def parse_share_link(url: str):
    """Return (workspace_id, view_id, token) from a sharing.clickup.com link.

    Example:
      https://sharing.clickup.com/90181817581/gr/h/2kzm2n7d-838/733f5c627a0c424
        workspace = 90181817581
        view_id   = 2kzm2n7d-838   (second-to-last path segment)
        token     = 733f5c627a0c424 (last path segment)
    """
    url = (url or "").strip()
    if not url:
        raise ScrapeError("กรุณาใส่ลิงค์ ClickUp public share")

    parsed = urlparse(url)
    if "clickup.com" not in (parsed.netloc or ""):
        raise ScrapeError("ลิงค์นี้ไม่ใช่ลิงค์ของ ClickUp")

    segments = [s for s in parsed.path.split("/") if s]
    if len(segments) < 3:
        raise ScrapeError(
            "รูปแบบลิงค์ไม่ถูกต้อง — ต้องเป็นลิงค์แชร์ View แบบสาธารณะ "
            "(เช่น .../<workspace>/gr/h/<viewId>/<token>)"
        )

    workspace = segments[0]
    if not re.fullmatch(r"\d+", workspace):
        raise ScrapeError("ไม่พบ workspace id ในลิงค์")

    view_id = segments[-2]
    token = segments[-1]
    return workspace, view_id, token


class ClickUpShareClient:
    def __init__(self, workspace: str, view_id: str, token: str,
                 timeout: int = 30, max_workers: int = 6):
        self.workspace = workspace
        self.view_id = view_id
        self.token = token
        self.timeout = timeout
        self.max_workers = max_workers
        self.session = requests.Session()
        self.session.headers.update(BASE_HEADERS)
        self.frontdoor: str | None = None

    # -- low level ---------------------------------------------------------
    def _handshake(self) -> str:
        r = self.session.get(HANDSHAKE_URL.format(ws=self.workspace),
                             timeout=self.timeout)
        if r.status_code != 200:
            raise ScrapeError(
                f"Handshake ล้มเหลว (HTTP {r.status_code}) — ตรวจสอบว่าลิงค์ยัง public อยู่")
        env = r.json().get("appEnvironment", {})
        api_url = env.get("apiUrl") or env.get("apiUrlV2")
        if not api_url:
            raise ScrapeError("ไม่พบ API endpoint จาก handshake")
        # apiUrl looks like https://frontdoor-...clickup.com/v1 -> strip version
        self.frontdoor = re.sub(r"/v\d+$", "", api_url)
        return self.frontdoor

    def _view_base(self) -> str:
        return f"{self.frontdoor}/view/v1/{self.workspace}/public/view/{self.view_id}"

    # -- public surface ----------------------------------------------------
    def fetch_view(self) -> dict:
        """Return the raw view payload (contains task ids + view metadata)."""
        if not self.frontdoor:
            self._handshake()
        url = f"{self._view_base()}?token={self.token}"
        r = self.session.get(url, timeout=self.timeout)
        if r.status_code == 404:
            raise ScrapeError("ไม่พบ View นี้ หรือลิงค์หมดอายุ/ถูกปิดการแชร์")
        if r.status_code != 200:
            raise ScrapeError(f"ดึงข้อมูล View ไม่สำเร็จ (HTTP {r.status_code})")
        return r.json()

    @staticmethod
    def task_ids_from_view(view: dict) -> list[str]:
        ids: list[str] = []
        seen = set()
        for group in (view.get("table", {}) or {}).get("groups", []) or []:
            for tid in group.get("task_ids", []) or []:
                if tid not in seen:
                    seen.add(tid)
                    ids.append(tid)
        return ids

    def fetch_task(self, task_id: str) -> dict:
        url = f"{self._view_base()}/task/{task_id}?token={self.token}"
        last_err = None
        for attempt in range(3):
            try:
                r = self.session.get(url, timeout=self.timeout)
                if r.status_code == 200:
                    return r.json()
                last_err = f"HTTP {r.status_code}"
            except requests.RequestException as exc:
                last_err = str(exc)
            time.sleep(0.6 * (attempt + 1))
        raise ScrapeError(f"ดึง task {task_id} ไม่สำเร็จ ({last_err})")


def scrape(url: str, progress=None) -> dict:
    """Scrape a public share link.

    Returns:
      {
        "view": <raw view json>,
        "view_name": str,
        "workspace": str,
        "tasks": [ <full task detail json>, ... ],   # in view order
        "errors": [ {task_id, error}, ... ],
      }
    progress(done, total, message) is called as work proceeds (optional).
    """
    def emit(done, total, msg):
        if progress:
            progress(done, total, msg)

    workspace, view_id, token = parse_share_link(url)
    client = ClickUpShareClient(workspace, view_id, token)

    emit(0, 0, "กำลังเชื่อมต่อ ClickUp...")
    view = client.fetch_view()
    task_ids = client.task_ids_from_view(view)
    total = len(task_ids)
    view_name = (view.get("view", {}) or {}).get("name") or "ClickUp View"
    emit(0, total, f"พบ {total} tasks ใน '{view_name}' — กำลังดึงรายละเอียด...")

    tasks: dict[str, dict] = {}
    errors = []
    done = 0
    with ThreadPoolExecutor(max_workers=client.max_workers) as pool:
        futures = {pool.submit(client.fetch_task, tid): tid for tid in task_ids}
        for fut in as_completed(futures):
            tid = futures[fut]
            try:
                tasks[tid] = fut.result()
            except ScrapeError as exc:
                errors.append({"task_id": tid, "error": str(exc)})
            done += 1
            emit(done, total, f"ดึงแล้ว {done}/{total} tasks")

    # preserve original view order
    ordered = [tasks[tid] for tid in task_ids if tid in tasks]
    return {
        "view": view,
        "view_name": view_name,
        "workspace": workspace,
        "tasks": ordered,
        "errors": errors,
    }


if __name__ == "__main__":
    import json
    import sys

    link = sys.argv[1] if len(sys.argv) > 1 else \
        "https://sharing.clickup.com/90181817581/gr/h/2kzm2n7d-838/733f5c627a0c424"
    result = scrape(link, progress=lambda d, t, m: print(f"  [{d}/{t}] {m}"))
    print(f"\nview: {result['view_name']}  tasks: {len(result['tasks'])}  "
          f"errors: {len(result['errors'])}")
    if result["tasks"]:
        t0 = result["tasks"][0]
        print("first task:", t0.get("name"),
              "| comments:", len(t0.get("comments", [])),
              "| subtasks:", len(t0.get("subtasks", [])))
