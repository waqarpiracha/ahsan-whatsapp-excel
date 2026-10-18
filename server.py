"""
Ahsan WhatsApp -> Excel importer (v1).

Ahsan exports the "EV charging solutions by Comet" group chat from
WhatsApp, uploads the .txt here TOGETHER WITH his current EV Master
Sheet (.xlsx), picks the export's date format, and gets back his sheet
with new survey rows appended (plus a Review sheet for the ones that
need his eye).

Endpoints:
  GET  /          -> the web app (app/index.html)
  POST /process   -> multipart: master (.xlsx) + chat (.txt) + date_format
                     -> output/*.xlsx ; JSON summary
  GET  /download  -> the finished .xlsx

No auth, no database. Flat files only (pilot scale).
Run:  python server.py  ->  http://localhost:5000
"""

import os
import shutil
import tempfile
from datetime import date
from flask import Flask, request, jsonify, send_file, send_from_directory

from parse_chat import parse_export, extract_rows
from write_workbook import append_rows

BASE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE, "app")
OUTPUT_ROOT = os.path.join(BASE, "output")
os.makedirs(OUTPUT_ROOT, exist_ok=True)

app = Flask(__name__, static_folder=APP_DIR, static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB


def _ext(name):
    return os.path.splitext(name or "")[1].lower()


@app.get("/")
def index():
    return send_from_directory(APP_DIR, "index.html")


@app.post("/process")
def process():
    master = request.files.get("master")
    chat = request.files.get("chat")
    date_format = (request.form.get("date_format") or "DMY").upper()
    if date_format not in ("DMY", "MDY"):
        return jsonify({"ok": False,
                        "error": "date_format must be DMY or MDY"}), 400
    if not master or _ext(master.filename) != ".xlsx":
        return jsonify({"ok": False,
                        "error": "upload your current EV Master Sheet (.xlsx)"}), 400
    if not chat or _ext(chat.filename) != ".txt":
        return jsonify({"ok": False,
                        "error": "upload the WhatsApp chat export (.txt)"}), 400

    tmpdir = tempfile.mkdtemp(prefix="ahsan_")
    try:
        master_path = os.path.join(tmpdir, "master.xlsx")
        chat_path = os.path.join(tmpdir, "chat.txt")
        master.save(master_path)
        chat.save(chat_path)

        try:
            messages = parse_export(chat_path, date_format=date_format)
        except Exception as e:
            return jsonify({"ok": False,
                            "error": f"could not read chat export: {e}"}), 400
        rows, review, stats = extract_rows(messages)

        out_name = f"EV_Master_{date.today().isoformat()}.xlsx"
        out_path = os.path.join(OUTPUT_ROOT, out_name)
        try:
            written, n_review = append_rows(master_path, rows, review,
                                            out_path,
                                            date_format=date_format)
        except ValueError as e:
            return jsonify({"ok": False, "error": str(e)}), 400

        return jsonify({
            "ok": True,
            "messages": stats["total"],
            "rows": written,
            "by_type": {k: stats[k] for k in
                        ("survey", "assign", "install", "cable")},
            "review": n_review,
            "ignored": stats["ignored"],
            "download": "/download?file=" + out_name,
        })
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@app.get("/status")
def status():
    """Lightweight ping for the ↻ refresh button (wakes the free-tier server)."""
    return jsonify({"ok": True, "service": "ahsan-whatsapp-excel"})


@app.get("/download")
def download():
    name = os.path.basename(request.args.get("file", ""))
    path = os.path.join(OUTPUT_ROOT, name)
    if not name.endswith(".xlsx") or not os.path.exists(path):
        return jsonify({"ok": False, "error": "not processed yet"}), 404
    return send_file(path, as_attachment=True, download_name=name)


if __name__ == "__main__":
    print("Ahsan WhatsApp->Excel  |  http://localhost:5000")
    app.run(host="0.0.0.0", port=5000)
