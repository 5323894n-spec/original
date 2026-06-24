from __future__ import annotations

import io
import tempfile
import uuid
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file
from werkzeug.utils import secure_filename

from app import OUTPUT_NAME, build_workbook, process_rows, read_source, validate_template


BASE_DIR = Path(__file__).resolve().parent
ALLOWED_EXTENSIONS = {".xlsx", ".csv"}
MAX_UPLOAD_SIZE = 50 * 1024 * 1024

web = Flask(__name__)
web.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE

# Готовые отчеты хранятся только в памяти до перезапуска приложения.
reports: dict[str, bytes] = {}


def validate_uploaded_file(upload, label: str) -> str:
    if upload is None or not upload.filename:
        raise ValueError(f"Не выбран файл: {label}.")
    suffix = Path(upload.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError(f"{label}: поддерживаются только .xlsx и .csv.")
    return upload.filename


@web.get("/")
def index():
    return render_template("index.html")


@web.get("/health")
def health():
    return {"status": "ok"}


@web.post("/process")
def process():
    source_upload = request.files.get("source")
    template_upload = request.files.get("template")

    try:
        source_name = validate_uploaded_file(source_upload, "исходная таблица")
        template_name = validate_uploaded_file(template_upload, "шаблон")

        with tempfile.TemporaryDirectory(prefix="trip_report_") as temp_dir:
            temp_path = Path(temp_dir)
            source_path = temp_path / (
                "source" + Path(secure_filename(source_name)).suffix.lower()
            )
            template_path = temp_path / (
                "template" + Path(secure_filename(template_name)).suffix.lower()
            )
            source_upload.save(source_path)
            template_upload.save(template_path)

            source = read_source(str(source_path))
            source.source_name = source_name
            validate_template(str(template_path))
            confirmation, errors, counters = process_rows(source)
            workbook = build_workbook(
                confirmation,
                errors,
                counters,
                source_name,
                template_name,
            )

            output = io.BytesIO()
            workbook.save(output)
            token = uuid.uuid4().hex
            reports[token] = output.getvalue()

            # Не даем локальному процессу бесконечно накапливать отчеты.
            while len(reports) > 10:
                reports.pop(next(iter(reports)))

        return render_template(
            "index.html",
            success=True,
            token=token,
            counters=counters,
            error_count=len(errors),
            source_name=source_name,
        )
    except Exception as exc:
        return render_template("index.html", error=str(exc)), 400


@web.get("/download/<token>")
def download(token: str):
    content = reports.get(token)
    if content is None:
        abort(404)
    return send_file(
        io.BytesIO(content),
        as_attachment=True,
        download_name=OUTPUT_NAME,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@web.errorhandler(413)
def file_too_large(_error):
    return render_template(
        "index.html",
        error="Размер загружаемых файлов превышает 50 МБ.",
    ), 413


if __name__ == "__main__":
    web.run(host="127.0.0.1", port=5000, debug=False)
