"""Фабрика приложения Flask."""
from flask import Flask

from .config import Config


def create_app(config_class: type = Config) -> Flask:
    app = Flask(__name__)
    app.config.from_object(config_class)
    app.config["MAX_CONTENT_LENGTH"] = config_class.MAX_CONTENT_LENGTH

    # Папки instance/uploads/outputs
    config_class.ensure_dirs()

    # БД
    from . import db
    db.init_app(app)
    with app.app_context():
        db.init_db()

    # Блюпринты
    from . import auth, main
    app.register_blueprint(auth.bp)
    app.register_blueprint(main.bp)

    return app
