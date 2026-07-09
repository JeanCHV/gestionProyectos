from __future__ import annotations

from main import app
from db import reset_db


if __name__ == "__main__":
    with app.app_context():
        reset_db()
    print("SQLite database reset.")
