from __future__ import annotations

import json
import os
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory
from waitress import serve

from office_convert import app_temp_dir, convert_office_to_pdf, unique_pdf_path
from pdf_ops import PdfLibrary, describe_doc, export_pages, split_ranges


APP_NAME = "PDF Editor-lite"
HOST = "127.0.0.1"
PORT = 8765


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


app = Flask(__name__, static_folder=str(resource_path("static")), static_url_path="")
library = PdfLibrary()
dialog_lock = threading.Lock()


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.post("/api/open-files")
def open_files():
    try:
        paths = ask_open_files()
        docs = library.add_paths(paths) if paths else []
        return jsonify({"cancelled": not bool(paths), "documents": [describe_doc(doc) for doc in docs]})
    except Exception as exc:
        return error_response(exc)


@app.post("/api/convert-office-add")
def convert_office_add():
    try:
        paths = ask_office_files()
        if not paths:
            return jsonify({"cancelled": True})

        converted = []
        pdf_paths = []
        target_dir = app_temp_dir() / time.strftime("%Y%m%d-%H%M%S")
        for source in paths:
            output = unique_pdf_path(target_dir, Path(source).name)
            result = convert_office_to_pdf(source, output)
            converted.append(result)
            pdf_paths.append(result["path"])

        docs = library.add_paths(pdf_paths)
        return jsonify({
            "cancelled": False,
            "converted": converted,
            "documents": [describe_doc(doc) for doc in docs],
        })
    except Exception as exc:
        write_error_log(exc)
        return error_response(exc)


@app.post("/api/convert-office-save")
def convert_office_save():
    try:
        paths = ask_office_files()
        if not paths:
            return jsonify({"cancelled": True})

        if len(paths) == 1:
            output = ask_save_file(f"{Path(paths[0]).stem}.pdf")
            if not output:
                return jsonify({"cancelled": True})
            results = [convert_office_to_pdf(paths[0], output)]
        else:
            output_dir = ask_directory("选择转换后 PDF 的保存文件夹")
            if not output_dir:
                return jsonify({"cancelled": True})
            results = []
            for source in paths:
                output = unique_pdf_path(output_dir, Path(source).name)
                results.append(convert_office_to_pdf(source, output))

        return jsonify({"cancelled": False, "results": results})
    except Exception as exc:
        write_error_log(exc)
        return error_response(exc)


@app.get("/api/file/<doc_id>")
def get_file(doc_id: str):
    try:
        doc = library.get(doc_id)
        return send_file(doc.path, mimetype="application/pdf", as_attachment=False, download_name=doc.name)
    except Exception as exc:
        return error_response(exc, status=404)


@app.post("/api/export")
def export_pdf():
    try:
        payload = request.get_json(force=True)
        page_items = payload.get("pages", [])
        default_name = payload.get("defaultName") or "输出.pdf"
        output = ask_save_file(default_name)
        if not output:
            return jsonify({"cancelled": True})
        result = export_pages(library, page_items, output)
        return jsonify({"cancelled": False, "result": result})
    except Exception as exc:
        write_error_log(exc)
        return error_response(exc)


@app.post("/api/split")
def split_pdf():
    try:
        payload = request.get_json(force=True)
        ranges = payload.get("ranges", [])
        output_dir = ask_directory("选择拆分输出文件夹")
        if not output_dir:
            return jsonify({"cancelled": True})
        results = split_ranges(library, ranges, output_dir)
        return jsonify({"cancelled": False, "results": results})
    except Exception as exc:
        write_error_log(exc)
        return error_response(exc)


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "name": APP_NAME})


def ask_open_files() -> tuple[str, ...]:
    with dialog_lock:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            paths = filedialog.askopenfilenames(
                title="选择 PDF 文件",
                filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
            )
            return tuple(paths)
        finally:
            root.destroy()


def ask_office_files() -> tuple[str, ...]:
    with dialog_lock:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            paths = filedialog.askopenfilenames(
                title="选择 Word 或 PowerPoint 文件",
                filetypes=[
                    ("Office 文件", "*.doc *.docx *.ppt *.pptx"),
                    ("Word 文件", "*.doc *.docx"),
                    ("PowerPoint 文件", "*.ppt *.pptx"),
                    ("所有文件", "*.*"),
                ],
            )
            return tuple(paths)
        finally:
            root.destroy()


def ask_save_file(default_name: str) -> str:
    with dialog_lock:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            return filedialog.asksaveasfilename(
                title="保存 PDF",
                defaultextension=".pdf",
                initialfile=default_name,
                filetypes=[("PDF 文件", "*.pdf")],
                confirmoverwrite=True,
            )
        finally:
            root.destroy()


def ask_directory(title: str) -> str:
    with dialog_lock:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            return filedialog.askdirectory(title=title)
        finally:
            root.destroy()


def error_response(exc: Exception, status: int = 400):
    return jsonify({"error": str(exc)}), status


def write_error_log(exc: Exception) -> None:
    log_dir = Path(os.environ.get("LOCALAPPDATA", Path.cwd())) / APP_NAME
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "error.log"
    payload = {
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "error": str(exc),
        "traceback": traceback.format_exc(),
    }
    with log_path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(payload, ensure_ascii=False) + "\n")


def open_browser() -> None:
    time.sleep(0.7)
    webbrowser.open(f"http://{HOST}:{PORT}/")


def main() -> None:
    if os.environ.get("PDF_TOOL_NO_BROWSER") != "1":
        threading.Thread(target=open_browser, daemon=True).start()
    serve(app, host=HOST, port=PORT, threads=8)


if __name__ == "__main__":
    main()
