"""Авторизация: регистрация, вход, выход. Простая, без верификации e-mail."""
import functools

from flask import (
    Blueprint, flash, g, redirect, render_template, request, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash

from .db import get_db

bp = Blueprint("auth", __name__, url_prefix="/auth")


@bp.before_app_request
def load_logged_in_user():
    """Подгружает текущего пользователя в g.user для каждого запроса."""
    user_id = session.get("user_id")
    if user_id is None:
        g.user = None
    else:
        g.user = get_db().execute(
            "SELECT * FROM users WHERE id = ?", (user_id,)
        ).fetchone()


def login_required(view):
    """Декоратор: пускает только авторизованных пользователей."""
    @functools.wraps(view)
    def wrapped_view(**kwargs):
        if g.user is None:
            return redirect(url_for("auth.login"))
        return view(**kwargs)
    return wrapped_view


@bp.route("/register", methods=("GET", "POST"))
def register():
    if g.user:
        return redirect(url_for("main.index"))
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        password2 = request.form.get("password2") or ""

        error = None
        if not username:
            error = "Введите имя пользователя."
        elif len(password) < 4:
            error = "Пароль должен быть не короче 4 символов."
        elif password != password2:
            error = "Пароли не совпадают."

        if error is None:
            db = get_db()
            try:
                db.execute(
                    "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                    (username, generate_password_hash(password)),
                )
                db.commit()
            except db.IntegrityError:
                error = f"Пользователь «{username}» уже существует."
            else:
                # сразу логиним
                user = db.execute(
                    "SELECT * FROM users WHERE username = ?", (username,)
                ).fetchone()
                session.clear()
                session["user_id"] = user["id"]
                return redirect(url_for("main.index"))

        flash(error, "error")

    return render_template("register.html")


@bp.route("/login", methods=("GET", "POST"))
def login():
    if g.user:
        return redirect(url_for("main.index"))
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()

        error = None
        if user is None or not check_password_hash(user["password_hash"], password):
            error = "Неверное имя пользователя или пароль."

        if error is None:
            session.clear()
            session["user_id"] = user["id"]
            return redirect(url_for("main.index"))

        flash(error, "error")

    return render_template("login.html")


@bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))
