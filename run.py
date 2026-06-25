"""Точка входа для локального запуска.

    python3 run.py
или
    flask --app run run --debug
"""
from app import create_app

app = create_app()

if __name__ == "__main__":
    # host=0.0.0.0 — чтобы открывалось и с других устройств в локальной сети
    app.run(host="127.0.0.1", port=8000, debug=True)
