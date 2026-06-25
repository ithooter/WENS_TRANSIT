"""Конфигурация приложения."""
import os
from pathlib import Path

# Корень проекта (на уровень выше пакета app/)
BASE_DIR = Path(__file__).resolve().parent.parent
INSTANCE_DIR = BASE_DIR / "instance"        # БД, загрузки, результаты (в .gitignore)
ENGINES_DIR = Path(__file__).resolve().parent / "engines"


class Config:
    # ВНИМАНИЕ: для продакшна задай переменную окружения SECRET_KEY.
    SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-change-me")

    INSTANCE_DIR = INSTANCE_DIR
    DB_PATH = INSTANCE_DIR / "app.db"
    UPLOAD_DIR = INSTANCE_DIR / "uploads"
    OUTPUT_DIR = INSTANCE_DIR / "outputs"

    ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsb"}
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024     # 50 МБ на файл

    @staticmethod
    def ensure_dirs():
        for d in (INSTANCE_DIR, Config.UPLOAD_DIR, Config.OUTPUT_DIR):
            Path(d).mkdir(parents=True, exist_ok=True)
