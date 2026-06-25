"""Основные маршруты: загрузка → выбор документов → генерация → скачивание."""
import uuid
from pathlib import Path

from flask import (
    Blueprint, abort, current_app, flash, g, redirect,
    render_template, request, send_file, url_for
)
from werkzeug.utils import secure_filename

from .auth import login_required
from .db import get_db
from . import service

bp = Blueprint("main", __name__)


def _allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in current_app.config["ALLOWED_EXTENSIONS"]


@bp.route("/")
@login_required
def index():
    return render_template("index.html")


@bp.route("/upload", methods=["POST"])
@login_required
def upload():
    file = request.files.get("source")
    if not file or file.filename == "":
        flash("Выберите файл-исходник.", "error")
        return redirect(url_for("main.index"))
    if not _allowed(file.filename):
        flash("Поддерживаются только файлы .xlsx, .xls, .xlsb", "error")
        return redirect(url_for("main.index"))

    job_id = uuid.uuid4().hex
    job_dir = Path(current_app.config["UPLOAD_DIR"]) / str(g.user["id"]) / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    # Сохраняем расширение явно (secure_filename может «съесть» кириллицу/точку)
    ext = Path(file.filename).suffix.lower()
    base = secure_filename(Path(file.filename).stem) or "source"
    filename = f"{base}{ext}"
    source_path = job_dir / filename
    file.save(source_path)

    db = get_db()
    db.execute(
        "INSERT INTO jobs (id, user_id, source_name, source_path) VALUES (?, ?, ?, ?)",
        (job_id, g.user["id"], filename, str(source_path)),
    )
    db.commit()

    return redirect(url_for("main.options", job_id=job_id))


def _get_job(job_id: str):
    job = get_db().execute(
        "SELECT * FROM jobs WHERE id = ? AND user_id = ?",
        (job_id, g.user["id"]),
    ).fetchone()
    if job is None:
        abort(404)
    return job


@bp.route("/options/<job_id>")
@login_required
def options(job_id):
    job = _get_job(job_id)
    return render_template(
        "options.html",
        job=job,
        output_types=service.OUTPUT_TYPES,
        parties=service.parties_for_ui(),
    )


@bp.route("/generate/<job_id>", methods=["POST"])
@login_required
def generate(job_id):
    job = _get_job(job_id)
    selections = request.form.getlist("outputs")
    if not selections:
        flash("Отметьте хотя бы один тип документа.", "error")
        return redirect(url_for("main.options", job_id=job_id))

    out_dir = Path(current_app.config["OUTPUT_DIR"]) / str(g.user["id"]) / job_id
    result = service.generate_documents(
        source_path=job["source_path"],
        selections=selections,
        sender_text=request.form.get("sender"),
        receiver_text=request.form.get("receiver"),
        out_dir=out_dir,
    )

    files = [{"name": p.name} for p in result.created]
    return render_template(
        "result.html",
        job=job,
        files=files,
        messages=result.messages,
    )


@bp.route("/download/<job_id>/<path:filename>")
@login_required
def download(job_id, filename):
    _get_job(job_id)  # проверка владельца
    out_dir = Path(current_app.config["OUTPUT_DIR"]) / str(g.user["id"]) / job_id
    safe_name = secure_filename(filename)
    target = (out_dir / safe_name).resolve()
    # защита от выхода за пределы папки задания
    if out_dir.resolve() not in target.parents or not target.exists():
        abort(404)
    return send_file(target, as_attachment=True, download_name=safe_name)
