"""Local web app: upload a PDF, download an editable Keynote (.key).

Run `pdf2key-web` (opens http://127.0.0.1:5001). Needs macOS with Keynote.
"""
import argparse
import tempfile
import threading
import uuid
import webbrowser
from importlib import resources
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from . import converter
from . import report as reporting

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024  # 1 GB

WORK_DIR = Path(tempfile.mkdtemp(prefix="pdf2key_"))
jobs = {}
keynote_lock = threading.Lock()  # Keynote handles one import at a time


def run_job(job_id):
    job = jobs[job_id]

    def progress(message, fraction):
        job.update(message=message, progress=fraction)

    try:
        is_pdf = job["pdf"].suffix.lower() == ".pdf"
        rep = reporting.new("PDF" if is_pdf else "HTML deck")
        job["message"] = "Waiting for Keynote"
        if is_pdf:
            job["missing_fonts"] = converter.missing_fonts(job["pdf"])
            with keynote_lock:
                reports = converter.convert(
                    job["pdf"], job["key"], separate_images=job["separate_images"], verify=True,
                    progress=progress, preview_dir=job["dir"] / "preview", report=rep)
        else:  # HTML deck export (.zip / .html): experimental, movable pictures
            from .html import convert_html
            with keynote_lock:
                reports = convert_html(job["pdf"], job["key"], verify=True, progress=progress,
                                       preview_dir=job["dir"] / "preview", report=rep, deck=job.get("deck"))
        job["identified_fonts"] = rep["identified"]
        job["report"] = rep
        job["reports"] = [r.__dict__ for r in reports]
        job.update(state="done", message="Done", progress=1)
    except Exception as e:
        job.update(state="error", error=str(e))


@app.post("/convert")
def convert():
    f = request.files.get("pdf")
    # an uploaded .html would arrive without the images and styles beside it,
    # so HTML decks come in as the exported .zip
    if not f or not f.filename.lower().endswith((".pdf", ".zip")):
        return jsonify(error="Please upload a PDF, or an HTML deck export as its .zip."), 400
    job_id = uuid.uuid4().hex
    job_dir = WORK_DIR / job_id
    job_dir.mkdir()
    stem = Path(secure_filename(f.filename)).stem or "presentation"
    pdf = job_dir / f"{stem}{Path(f.filename).suffix.lower()}"
    f.save(pdf)
    jobs[job_id] = dict(
        state="running", message="Starting", progress=0, pdf=pdf, dir=job_dir,
        key=job_dir / f"{stem}.key", name=f"{stem}.key",
        separate_images=request.form.get("separate_images") == "1", reports=[], missing_fonts=[], identified_fonts=[])
    if pdf.suffix == ".zip":
        from .html import decks_in
        try:
            decks = decks_in(pdf)
        except Exception as e:
            return jsonify(error=f"Couldn't read the export: {e}"), 400
        if not decks:
            return jsonify(error="No HTML deck (a page with <section> slides) in this .zip."), 400
        if len(decks) > 1:  # the export holds several decks: let the person pick
            jobs[job_id]["state"] = "choose"
            return jsonify(id=job_id, choices=[{"name": d, "slides": n} for d, n in decks])
    threading.Thread(target=run_job, args=(job_id,), daemon=True).start()
    return jsonify(id=job_id)


@app.post("/start/<job_id>")
def start(job_id):
    job = jobs.get(job_id)
    if not job or job["state"] != "choose":
        return jsonify(error="Unknown job"), 404
    job.update(state="running", deck=(request.get_json(silent=True) or {}).get("deck"))
    threading.Thread(target=run_job, args=(job_id,), daemon=True).start()
    return jsonify(id=job_id)


@app.get("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="Unknown job"), 404
    keys = ("state", "message", "progress", "error", "reports", "missing_fonts", "identified_fonts", "name", "report")
    return jsonify({k: job.get(k) for k in keys})


@app.get("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job["state"] != "done":
        return "Not ready", 404
    return send_file(job["key"], as_attachment=True, download_name=job["name"])


@app.get("/preview/<job_id>/<name>")
def preview(job_id, name):
    job = jobs.get(job_id)
    path = job["dir"] / "preview" / secure_filename(name) if job else None
    if not path or not path.exists():
        return "Not found", 404
    return send_file(path)


@app.get("/")
def index():
    return resources.files("pdf2key").joinpath("static/index.html").read_text(encoding="utf-8")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="pdf2key-web", description="Run the PDF to Keynote web app.")
    ap.add_argument("--port", type=int, default=5001)
    ap.add_argument("--no-browser", action="store_true", help="don't open a browser tab")
    args = ap.parse_args(argv)
    if not args.no_browser:
        threading.Timer(1, lambda: webbrowser.open(f"http://127.0.0.1:{args.port}")).start()
    app.run(host="127.0.0.1", port=args.port)


if __name__ == "__main__":
    main()
