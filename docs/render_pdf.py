"""Render manual.html -> คู่มือการใช้งาน.pdf via headless Chromium (perfect Thai)."""
import os
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "manual.html")
PDF = os.path.join(HERE, "คู่มือการใช้งาน ClickUp-Excel.pdf")

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto("file:///" + HTML.replace("\\", "/"), wait_until="networkidle")
    page.pdf(
        path=PDF,
        format="A4",
        print_background=True,
        margin={"top": "0", "bottom": "0", "left": "0", "right": "0"},
    )
    browser.close()

print("PDF created:", PDF, "| size:", os.path.getsize(PDF), "bytes")
