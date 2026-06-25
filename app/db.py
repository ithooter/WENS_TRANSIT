"""Минимальный слой работы с SQLite (без ORM)."""
import sqlite3
from flask import g, current_app


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DB_PATH"])
        g.db.row_factory = sqlite3.Row
    return g.db


def close_db(e=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """Создаёт таблицы, если их нет. Вызывается при старте приложения."""
    db = sqlite3.connect(current_app.config["DB_PATH"])
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at    TEXT NOT NULL DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS jobs (
            id          TEXT PRIMARY KEY,
            user_id     INTEGER NOT NULL,
            source_name TEXT NOT NULL,
            source_path TEXT NOT NULL,
            created_at  TEXT NOT NULL DEFAULT (datetime('now')),
            FOREIGN KEY (user_id) REFERENCES users(id)
        );
        """
    )
    db.commit()
    db.close()


def init_app(app):
    app.teardown_appcontext(close_db)
