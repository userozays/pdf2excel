"""Yerel web arayüzü: http://127.0.0.1:5050  (dosyalar bilgisayardan çıkmaz)"""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path

import pypdfium2 as pdfium
from flask import Flask, abort, jsonify, render_template, request, send_file

from pdf2excel import Options, convert_file

# PyInstaller ile paketlenince şablonlar geçici açılım klasöründe olur
BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
app = Flask(__name__, template_folder=str(BASE / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024
JOBS = Path(tempfile.gettempdir()) / "pdf2excel_jobs"
JOBS.mkdir(exist_ok=True)
PREVIEW_ROWS = 500


def _cleanup(max_age=3600):
    now = time.time()
    for d in JOBS.iterdir():
        if d.is_dir() and now - d.stat().st_mtime > max_age:
            shutil.rmtree(d, ignore_errors=True)


def _job(job_id: str) -> Path:
    try:
        uuid.UUID(job_id)
    except ValueError:
        abort(404)
    d = JOBS / job_id
    if not d.exists():
        abort(404)
    return d


SAMPLES = Path(__file__).parent / "tests" / "out"


def _samples():
    return sorted(SAMPLES.glob("*.pdf")) if SAMPLES.exists() else []


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/samples")
def api_samples():
    return jsonify([p.name for p in _samples()])


@app.post("/api/convert")
def api_convert():
    _cleanup()
    f = request.files.get("pdf")
    sample = request.form.get("sample")
    if sample:
        src = next((p for p in _samples() if p.name == sample), None)
        if src is None:
            return jsonify(error="Örnek bulunamadı."), 404
    elif not f or not f.filename.lower().endswith(".pdf"):
        return jsonify(error="Lütfen bir PDF dosyası seçin."), 400
    job_id = str(uuid.uuid4())
    d = JOBS / job_id
    d.mkdir()
    pdf = d / "in.pdf"
    if sample:
        shutil.copyfile(src, pdf)
        name = src.stem
    else:
        f.save(pdf)
        name = Path(f.filename).stem
    (d / "name.txt").write_text(name, encoding="utf-8")
    opts = Options(
        mode=request.form.get("mode", "tek"),
        locale=request.form.get("locale", "auto"),
        convert=request.form.get("convert", "1") == "1",
        join_wrapped=request.form.get("join", "1") == "1",
        tr_fix=request.form.get("trfix", "auto"),
        skip_repeated=request.form.get("skiprep", "1") == "1",
    )
    try:
        res, rep = convert_file(str(pdf), str(d / "out.xlsx"), opts)
    except Exception as e:  # bozuk/şifreli PDF vb.
        shutil.rmtree(d, ignore_errors=True)
        app.logger.warning("PDF okunamadı: %s", e)
        msg = ("PDF şifreli; önce şifresini kaldırıp tekrar deneyin." if "password" in str(e).lower() or "encrypt" in str(e).lower()
               else "PDF okunamadı. Dosya bozuk ya da gerçek bir PDF değil.")
        return jsonify(error=msg), 422

    flagged = {(f["sheet"], f["row"], f["col"]): f["notes"] for f in rep["flags"]}
    sheets = []
    for s in res.sheets:
        ncols = max([c + 1 for r in s.rows for c in r.cells] + [c1 + 1 for *_, c1 in s.merges] + [1])
        rows = []
        for ri, r in enumerate(s.rows[:PREVIEW_ROWS]):
            cells = {}
            for c, t in r.cells.items():
                typed = s.typed.get((ri, c))
                kind = None if not typed else ("tarih" if hasattr(typed[0], "year") else "sayi")
                cells[c] = {"t": t, "k": kind, "n": flagged.get((s.name, ri, c), [])}
            rows.append({"p": r.page, "c": cells})
        sheets.append({"name": s.name, "ncols": ncols, "nrows": len(s.rows), "rows": rows,
                       "merges": [m for m in s.merges if m[0] < PREVIEW_ROWS]})
    with pdfium.PdfDocument(str(pdf)) as doc:
        npages = len(doc)
    return jsonify(id=job_id, report=rep, sheets=sheets, pages=npages)


@app.get("/api/download/<job_id>")
def api_download(job_id):
    d = _job(job_id)
    name = (d / "name.txt").read_text(encoding="utf-8")
    return send_file(d / "out.xlsx", as_attachment=True, download_name=f"{name}.xlsx")


@app.get("/api/page/<job_id>/<int:n>.png")
def api_page(job_id, n):
    d = _job(job_id)
    png = d / f"p{n}.png"
    if not png.exists():
        with pdfium.PdfDocument(str(d / "in.pdf")) as doc:
            if not 1 <= n <= len(doc):
                abort(404)
            doc[n - 1].render(scale=1.6).to_pil().save(png)
    return send_file(png, mimetype="image/png")


if __name__ == "__main__":
    import sys
    import webbrowser
    if "--no-browser" not in sys.argv:
        webbrowser.open("http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)
