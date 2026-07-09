from __future__ import annotations

import os
import sqlite3
from contextlib import closing
from datetime import datetime
from typing import Any

from flask import current_app, g


def get_db_path() -> str:
    return current_app.config["DATABASE"]


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        conn = sqlite3.connect(get_db_path())
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        g.db = conn
    return g.db


def close_db(_: Any = None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def query_all(sql: str, params: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
    rows = get_db().execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def query_one(sql: str, params: tuple[Any, ...] = ()) -> dict[str, Any] | None:
    row = get_db().execute(sql, params).fetchone()
    return dict(row) if row else None


def execute(sql: str, params: tuple[Any, ...] = ()) -> sqlite3.Cursor:
    conn = get_db()
    cursor = conn.execute(sql, params)
    conn.commit()
    return cursor


def make_code(prefix: str, pk: int) -> str:
    return f"{prefix}-{pk:04d}"


def iso_date(value: Any) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except TypeError:
            return str(value)
    return str(value)


def init_db() -> None:
    os.makedirs(os.path.dirname(get_db_path()), exist_ok=True)
    conn = sqlite3.connect(get_db_path())
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT NOT NULL,
            project_type TEXT,
            company TEXT,
            start_date TEXT,
            end_date TEXT,
            status TEXT DEFAULT 'Planificado',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS assets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            project_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            type TEXT,
            owner TEXT,
            value REAL DEFAULT 0,
            status TEXT DEFAULT 'Activo',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS risks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            project_id INTEGER NOT NULL,
            asset_id INTEGER,
            asset_name TEXT,
            name TEXT NOT NULL,
            description TEXT,
            cause TEXT,
            consequence TEXT,
            probability REAL DEFAULT 0,
            impact REAL DEFAULT 0,
            level TEXT,
            horizon TEXT,
            owner TEXT,
            status TEXT,
            strategy TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(asset_id) REFERENCES assets(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS mitigations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            project_id INTEGER NOT NULL,
            risk_id INTEGER,
            risk_name TEXT,
            preventive_action TEXT,
            corrective_action TEXT,
            owner TEXT,
            start_date TEXT,
            end_date TEXT,
            resources TEXT,
            status TEXT,
            evidence TEXT,
            strategy TEXT,
            progress INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(risk_id) REFERENCES risks(id) ON DELETE SET NULL
        );
        """
    )
    conn.commit()
    conn.close()


def reset_db() -> None:
    path = get_db_path()
    if os.path.exists(path):
        os.remove(path)
    init_db()


def list_projects() -> list[dict[str, Any]]:
    return query_all("SELECT * FROM projects ORDER BY id DESC")


def get_project(project_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM projects WHERE id = ?", (project_id,))


def create_project(data: dict[str, Any]) -> dict[str, Any]:
    cursor = execute(
        """
        INSERT INTO projects (code, name, project_type, company, start_date, end_date, status)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            data.get("name", "").strip(),
            data.get("project_type", "").strip(),
            data.get("company", "").strip(),
            iso_date(data.get("start_date")),
            iso_date(data.get("end_date")),
            data.get("status", "Planificado"),
        ),
    )
    project_id = cursor.lastrowid
    code = make_code("PRJ", project_id)
    execute("UPDATE projects SET code = ? WHERE id = ?", (code, project_id))
    return get_project(project_id)


def update_project(project_id: int, data: dict[str, Any]) -> None:
    execute(
        """
        UPDATE projects
        SET name = ?, project_type = ?, company = ?, start_date = ?, end_date = ?, status = ?
        WHERE id = ?
        """,
        (
            data.get("name", "").strip(),
            data.get("project_type", "").strip(),
            data.get("company", "").strip(),
            iso_date(data.get("start_date")),
            iso_date(data.get("end_date")),
            data.get("status", "Planificado"),
            project_id,
        ),
    )


def delete_project(project_id: int) -> None:
    execute("DELETE FROM projects WHERE id = ?", (project_id,))


def list_assets(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        "SELECT * FROM assets WHERE project_id = ? ORDER BY id DESC",
        (project_id,),
    )


def get_asset(asset_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM assets WHERE id = ?", (asset_id,))


def create_asset(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    cursor = execute(
        """
        INSERT INTO assets (code, project_id, name, type, owner, value, status)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            data.get("name", "").strip(),
            data.get("type", "").strip(),
            data.get("owner", "").strip(),
            float(data.get("value") or 0),
            data.get("status", "Activo"),
        ),
    )
    asset_id = cursor.lastrowid
    execute("UPDATE assets SET code = ? WHERE id = ?", (make_code("ACT", asset_id), asset_id))
    return get_asset(asset_id)


def update_asset(asset_id: int, data: dict[str, Any]) -> None:
    execute(
        """
        UPDATE assets
        SET name = ?, type = ?, owner = ?, value = ?, status = ?
        WHERE id = ?
        """,
        (
            data.get("name", "").strip(),
            data.get("type", "").strip(),
            data.get("owner", "").strip(),
            float(data.get("value") or 0),
            data.get("status", "Activo"),
            asset_id,
        ),
    )


def delete_asset(asset_id: int) -> None:
    execute("DELETE FROM assets WHERE id = ?", (asset_id,))


def list_assets_for_select(project_id: int | None) -> list[dict[str, Any]]:
    return query_all(
        "SELECT id, code, name FROM assets WHERE project_id = ? ORDER BY name",
        (project_id,),
    ) if project_id else []


def list_risks(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT r.*, a.name AS asset_label
        FROM risks r
        LEFT JOIN assets a ON a.id = r.asset_id
        WHERE r.project_id = ?
        ORDER BY r.id DESC
        """,
        (project_id,),
    )


def get_risk(risk_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM risks WHERE id = ?", (risk_id,))


def create_risk(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    asset_id = data.get("asset_id")
    asset = get_asset(int(asset_id)) if asset_id else None
    probability = float(data.get("probability") or 0)
    impact = float(data.get("impact") or 0)
    level = data.get("level") or calculate_level(probability, impact)
    cursor = execute(
        """
        INSERT INTO risks (
            code, project_id, asset_id, asset_name, name, description, cause, consequence,
            probability, impact, level, horizon, owner, status, strategy
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            int(asset_id) if asset_id else None,
            asset["name"] if asset else data.get("asset_name", "").strip(),
            data.get("name", "").strip(),
            data.get("description", "").strip(),
            data.get("cause", "").strip(),
            data.get("consequence", "").strip(),
            probability,
            impact,
            level,
            data.get("horizon", "").strip(),
            data.get("owner", "").strip(),
            data.get("status", "Identificado"),
            data.get("strategy", "Mitigar"),
        ),
    )
    risk_id = cursor.lastrowid
    execute("UPDATE risks SET code = ? WHERE id = ?", (make_code("R", risk_id), risk_id))
    return get_risk(risk_id)


def update_risk(risk_id: int, data: dict[str, Any]) -> None:
    asset_id = data.get("asset_id")
    asset = get_asset(int(asset_id)) if asset_id else None
    probability = float(data.get("probability") or 0)
    impact = float(data.get("impact") or 0)
    level = data.get("level") or calculate_level(probability, impact)
    execute(
        """
        UPDATE risks
        SET asset_id = ?, asset_name = ?, name = ?, description = ?, cause = ?, consequence = ?,
            probability = ?, impact = ?, level = ?, horizon = ?, owner = ?, status = ?, strategy = ?
        WHERE id = ?
        """,
        (
            int(asset_id) if asset_id else None,
            asset["name"] if asset else data.get("asset_name", "").strip(),
            data.get("name", "").strip(),
            data.get("description", "").strip(),
            data.get("cause", "").strip(),
            data.get("consequence", "").strip(),
            probability,
            impact,
            level,
            data.get("horizon", "").strip(),
            data.get("owner", "").strip(),
            data.get("status", "Identificado"),
            data.get("strategy", "Mitigar"),
            risk_id,
        ),
    )


def delete_risk(risk_id: int) -> None:
    execute("DELETE FROM risks WHERE id = ?", (risk_id,))


def list_risk_names(project_id: int | None) -> list[dict[str, Any]]:
    return query_all(
        "SELECT id, code, name FROM risks WHERE project_id = ? ORDER BY name",
        (project_id,),
    ) if project_id else []


def list_mitigations(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT m.*, r.name AS risk_label
        FROM mitigations m
        LEFT JOIN risks r ON r.id = m.risk_id
        WHERE m.project_id = ?
        ORDER BY m.id DESC
        """,
        (project_id,),
    )


def get_mitigation(mitigation_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM mitigations WHERE id = ?", (mitigation_id,))


def create_mitigation(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    risk_id = data.get("risk_id")
    risk = get_risk(int(risk_id)) if risk_id else None
    cursor = execute(
        """
        INSERT INTO mitigations (
            code, project_id, risk_id, risk_name, preventive_action, corrective_action,
            owner, start_date, end_date, resources, status, evidence, strategy, progress
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            int(risk_id) if risk_id else None,
            risk["name"] if risk else data.get("risk_name", "").strip(),
            data.get("preventive_action", "").strip(),
            data.get("corrective_action", "").strip(),
            data.get("owner", "").strip(),
            iso_date(data.get("start_date")),
            iso_date(data.get("end_date")),
            data.get("resources", "").strip(),
            data.get("status", "Planificado"),
            data.get("evidence", "").strip(),
            data.get("strategy", "Mitigar"),
            int(data.get("progress") or 0),
        ),
    )
    mitigation_id = cursor.lastrowid
    execute("UPDATE mitigations SET code = ? WHERE id = ?", (make_code("MG", mitigation_id), mitigation_id))
    return get_mitigation(mitigation_id)


def update_mitigation(mitigation_id: int, data: dict[str, Any]) -> None:
    risk_id = data.get("risk_id")
    risk = get_risk(int(risk_id)) if risk_id else None
    execute(
        """
        UPDATE mitigations
        SET risk_id = ?, risk_name = ?, preventive_action = ?, corrective_action = ?,
            owner = ?, start_date = ?, end_date = ?, resources = ?, status = ?, evidence = ?,
            strategy = ?, progress = ?
        WHERE id = ?
        """,
        (
            int(risk_id) if risk_id else None,
            risk["name"] if risk else data.get("risk_name", "").strip(),
            data.get("preventive_action", "").strip(),
            data.get("corrective_action", "").strip(),
            data.get("owner", "").strip(),
            iso_date(data.get("start_date")),
            iso_date(data.get("end_date")),
            data.get("resources", "").strip(),
            data.get("status", "Planificado"),
            data.get("evidence", "").strip(),
            data.get("strategy", "Mitigar"),
            int(data.get("progress") or 0),
            mitigation_id,
        ),
    )


def delete_mitigation(mitigation_id: int) -> None:
    execute("DELETE FROM mitigations WHERE id = ?", (mitigation_id,))


def calculate_level(probability: float, impact: float) -> str:
    score = probability * impact
    if score < 15:
        return "Bajo"
    if score < 40:
        return "Medio"
    if score < 70:
        return "Alto"
    return "Critico"


def report_summary(project_id: int | None) -> dict[str, Any]:
    if not project_id:
        return {"total": 0, "low": 0, "medium": 0, "high": 0, "critical": 0, "mitigated": 0, "progress": 0}
    rows = query_all("SELECT level, status FROM risks WHERE project_id = ?", (project_id,))
    counts = {"Bajo": 0, "Medio": 0, "Alto": 0, "Critico": 0}
    mitigated = 0
    for row in rows:
        counts[row["level"] or "Bajo"] = counts.get(row["level"] or "Bajo", 0) + 1
        if (row["status"] or "").lower() in {"mitigado", "cerrado"}:
            mitigated += 1
    total = len(rows)
    progress = round((mitigated / total) * 100) if total else 0
    return {
        "total": total,
        "low": counts["Bajo"],
        "medium": counts["Medio"],
        "high": counts["Alto"],
        "critical": counts["Critico"],
        "mitigated": mitigated,
        "progress": progress,
    }
