"""PDF2Excel masaüstü uygulaması: yerel sunucu + yerel pencere (Windows'ta WebView2).

Ağa çıkılmaz: sunucu yalnız 127.0.0.1'de, rastgele bir portta dinler. Bu oturumun geçici
dosyaları kendi klasöründe tutulur ve pencere kapanınca silinir.

Test için:  PDF2Excel.exe --convert girdi.pdf cikti.xlsx   (pencere açmadan çevirir)
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import sys
import tempfile
import threading
import time
from pathlib import Path

import webview
from werkzeug.serving import make_server

import app as webapp


class Api:
    """Sayfadaki JavaScript'in çağırabildiği masaüstü işlevleri."""

    def __init__(self):
        # "_" ile başlamalı: pywebview js_api nesnesinin açık alanlarını JavaScript'e
        # aktarmaya çalışır ve pencerenin iç .NET nesnelerine takılır.
        self._window = None

    def save_xlsx(self, job_id: str, name: str):
        src = Path(webapp.JOBS) / Path(job_id).name / "out.xlsx"
        if not src.exists():
            return {"ok": False, "error": "Dosya bulunamadı, PDF'i yeniden çevirin."}
        stem = Path(name).stem or "PDF2Excel"
        picked = self._window.create_file_dialog(
            webview.FileDialog.SAVE, save_filename=f"{stem}.xlsx", file_types=("Excel dosyası (*.xlsx)",))
        if not picked:
            return {"ok": False}
        target = picked if isinstance(picked, str) else picked[0]
        if not target.lower().endswith(".xlsx"):
            target += ".xlsx"
        shutil.copyfile(src, target)
        return {"ok": True, "path": target}

    def open_file(self, path: str):
        if os.path.exists(path):
            os.startfile(path)  # Excel'de aç


def run_cli(pdf: str, xlsx: str) -> int:
    from pdf2excel import convert_file
    _, rep = convert_file(pdf, xlsx)
    Path(xlsx).with_suffix(".json").write_text(
        json.dumps({k: rep[k] for k in ("ok", "cells", "info", "pages", "balance")}, ensure_ascii=False, default=str),
        encoding="utf-8")
    return 0 if rep["ok"] else 1


def main():
    # Konsolsuz (.exe / pythonw) çalışınca stdout/stderr yoktur; yazma girişimleri
    # sunucuyu kilitleyebilir. Çıktıyı geçici klasördeki bir günlük dosyasına yönlendir.
    if sys.stdout is None or sys.stderr is None:
        log = open(Path(tempfile.gettempdir()) / "PDF2Excel.log", "w", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or log
        sys.stderr = sys.stderr or log
    logging.getLogger("werkzeug").setLevel(logging.ERROR)

    if len(sys.argv) == 4 and sys.argv[1] == "--convert":
        sys.exit(run_cli(sys.argv[2], sys.argv[3]))

    # çökme ya da zorla kapatmadan kalmış, bir günden eski oturum klasörlerini temizle
    for old in Path(tempfile.gettempdir()).glob("pdf2excel_*"):
        try:
            if old.is_dir() and old.name != "pdf2excel_jobs" and time.time() - old.stat().st_mtime > 86400:
                shutil.rmtree(old, ignore_errors=True)
        except OSError:
            pass
    webapp.JOBS = Path(tempfile.mkdtemp(prefix="pdf2excel_"))
    srv = make_server("127.0.0.1", int(os.environ.get("PDF2EXCEL_PORT", 0)), webapp.app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    api = Api()
    api._window = webview.create_window(
        "PDF2Excel", f"http://127.0.0.1:{srv.server_port}/", js_api=api,
        width=1320, height=880, min_size=(900, 600),
        hidden=os.environ.get("PDF2EXCEL_HIDDEN") == "1")
    try:
        webview.start()
    finally:
        srv.shutdown()
        shutil.rmtree(webapp.JOBS, ignore_errors=True)


if __name__ == "__main__":
    main()
