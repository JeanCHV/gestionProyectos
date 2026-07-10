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

        CREATE TABLE IF NOT EXISTS project_roles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            acquisition_type TEXT,
            risk_participation TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS project_deliverables (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            deliverable TEXT NOT NULL,
            status TEXT DEFAULT 'Pendiente',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS project_threats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            project_id INTEGER NOT NULL,
            threat TEXT NOT NULL,
            affected_asset TEXT,
            example TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS project_safeguards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            project_id INTEGER NOT NULL,
            threat_id INTEGER,
            threat_name TEXT,
            safeguard TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(threat_id) REFERENCES project_threats(id) ON DELETE SET NULL
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
            trigger TEXT,
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
    mitigation_columns = {row[1] for row in conn.execute("PRAGMA table_info(mitigations)").fetchall()}
    if "trigger" not in mitigation_columns:
        conn.execute("ALTER TABLE mitigations ADD COLUMN trigger TEXT")
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


def list_project_deliverables(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT *
        FROM project_deliverables
        WHERE project_id = ?
        ORDER BY id ASC
        """,
        (project_id,),
    )


def create_project_deliverable(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    cursor = execute(
        """
        INSERT INTO project_deliverables (project_id, deliverable, status)
        VALUES (?, ?, ?)
        """,
        (
            project_id,
            data.get("deliverable", "").strip(),
            data.get("status", "Pendiente"),
        ),
    )
    return query_one("SELECT * FROM project_deliverables WHERE id = ?", (cursor.lastrowid,)) or {}


def get_project_deliverable(deliverable_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM project_deliverables WHERE id = ?", (deliverable_id,))


def update_project_deliverable(deliverable_id: int, data: dict[str, Any]) -> None:
    execute(
        """
        UPDATE project_deliverables
        SET deliverable = ?, status = ?
        WHERE id = ?
        """,
        (
            data.get("deliverable", "").strip(),
            data.get("status", "Pendiente"),
            deliverable_id,
        ),
    )


def delete_project_deliverable(deliverable_id: int) -> None:
    execute("DELETE FROM project_deliverables WHERE id = ?", (deliverable_id,))


def list_project_threats(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT *
        FROM project_threats
        WHERE project_id = ?
        ORDER BY id ASC
        """,
        (project_id,),
    )


def get_project_threat(threat_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM project_threats WHERE id = ?", (threat_id,))


def create_project_threat(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    cursor = execute(
        """
        INSERT INTO project_threats (code, project_id, threat, affected_asset, example)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            data.get("threat", "").strip(),
            data.get("affected_asset", "").strip(),
            data.get("example", "").strip(),
        ),
    )
    threat_id = cursor.lastrowid
    execute("UPDATE project_threats SET code = ? WHERE id = ?", (make_code("AM", threat_id), threat_id))
    return query_one("SELECT * FROM project_threats WHERE id = ?", (threat_id,)) or {}


def update_project_threat(threat_id: int, data: dict[str, Any]) -> None:
    execute(
        """
        UPDATE project_threats
        SET threat = ?, affected_asset = ?, example = ?
        WHERE id = ?
        """,
        (
            data.get("threat", "").strip(),
            data.get("affected_asset", "").strip(),
            data.get("example", "").strip(),
            threat_id,
        ),
    )


def delete_project_threat(threat_id: int) -> None:
    execute("DELETE FROM project_threats WHERE id = ?", (threat_id,))


def list_project_safeguards(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT s.*, t.code AS threat_code
        FROM project_safeguards s
        LEFT JOIN project_threats t ON t.id = s.threat_id
        WHERE s.project_id = ?
        ORDER BY s.id ASC
        """,
        (project_id,),
    )


def get_project_safeguard(safeguard_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM project_safeguards WHERE id = ?", (safeguard_id,))


def create_project_safeguard(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    threat_id = data.get("threat_id")
    threat = get_project_threat(int(threat_id)) if threat_id else None
    cursor = execute(
        """
        INSERT INTO project_safeguards (code, project_id, threat_id, threat_name, safeguard)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            int(threat_id) if threat_id else None,
            threat["threat"] if threat else data.get("threat_name", "").strip(),
            data.get("safeguard", "").strip(),
        ),
    )
    safeguard_id = cursor.lastrowid
    execute("UPDATE project_safeguards SET code = ? WHERE id = ?", (make_code("SG", safeguard_id), safeguard_id))
    return query_one("SELECT * FROM project_safeguards WHERE id = ?", (safeguard_id,)) or {}


def update_project_safeguard(safeguard_id: int, data: dict[str, Any]) -> None:
    threat_id = data.get("threat_id")
    threat = get_project_threat(int(threat_id)) if threat_id else None
    execute(
        """
        UPDATE project_safeguards
        SET threat_id = ?, threat_name = ?, safeguard = ?
        WHERE id = ?
        """,
        (
            int(threat_id) if threat_id else None,
            threat["threat"] if threat else data.get("threat_name", "").strip(),
            data.get("safeguard", "").strip(),
            safeguard_id,
        ),
    )


def delete_project_safeguard(safeguard_id: int) -> None:
    execute("DELETE FROM project_safeguards WHERE id = ?", (safeguard_id,))


def list_project_roles(project_id: int | None) -> list[dict[str, Any]]:
    if not project_id:
        return []
    return query_all(
        """
        SELECT *
        FROM project_roles
        WHERE project_id = ?
        ORDER BY id ASC
        """,
        (project_id,),
    )


def create_project_role(project_id: int, data: dict[str, Any]) -> dict[str, Any]:
    cursor = execute(
        """
        INSERT INTO project_roles (project_id, role, acquisition_type, risk_participation)
        VALUES (?, ?, ?, ?)
        """,
        (
            project_id,
            data.get("role", "").strip(),
            data.get("acquisition_type", "").strip(),
            data.get("risk_participation", "").strip(),
        ),
    )
    return query_one("SELECT * FROM project_roles WHERE id = ?", (cursor.lastrowid,)) or {}


def get_project_role(role_id: int) -> dict[str, Any] | None:
    return query_one("SELECT * FROM project_roles WHERE id = ?", (role_id,))


def update_project_role(role_id: int, data: dict[str, Any]) -> None:
    execute(
        """
        UPDATE project_roles
        SET role = ?, acquisition_type = ?, risk_participation = ?
        WHERE id = ?
        """,
        (
            data.get("role", "").strip(),
            data.get("acquisition_type", "").strip(),
            data.get("risk_participation", "").strip(),
            role_id,
        ),
    )


def delete_project_role(role_id: int) -> None:
    execute("DELETE FROM project_roles WHERE id = ?", (role_id,))


def _normalize_asset_value(value: Any) -> Any:
    if value in (None, ""):
        return ""
    text = str(value).strip()
    try:
        number = float(text)
    except ValueError:
        return text
    return int(number) if number.is_integer() else number


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
            _normalize_asset_value(data.get("value")),
            data.get("status", "Activo"),
        ),
    )
    asset_id = cursor.lastrowid
    execute("UPDATE assets SET code = ? WHERE id = ?", (make_code("ACT", asset_id), asset_id))
    return get_asset(asset_id)


def update_asset(asset_id: int, data: dict[str, Any]) -> None:
    current = get_asset(asset_id) or {}
    execute(
        """
        UPDATE assets
        SET name = ?, type = ?, owner = ?, value = ?, status = ?
        WHERE id = ?
        """,
        (
            data.get("name", current.get("name", "")).strip(),
            data.get("type", current.get("type", "")).strip(),
            data.get("owner", current.get("owner", "")).strip(),
            _normalize_asset_value(data.get("value", current.get("value", ""))),
            data.get("status", current.get("status", "Activo")),
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
    current = get_risk(risk_id) or {}
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
            data.get("name", current.get("name", "")).strip(),
            data.get("description", current.get("description", "")).strip(),
            data.get("cause", current.get("cause", "")).strip(),
            data.get("consequence", current.get("consequence", "")).strip(),
            probability,
            impact,
            level,
            data.get("horizon", current.get("horizon", "")).strip(),
            data.get("owner", current.get("owner", "")).strip(),
            data.get("status", current.get("status", "Identificado")),
            data.get("strategy", current.get("strategy", "Mitigar")),
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
            trigger, owner, start_date, end_date, resources, status, evidence, strategy, progress
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "TMP",
            project_id,
            int(risk_id) if risk_id else None,
            risk["name"] if risk else data.get("risk_name", "").strip(),
            data.get("preventive_action", "").strip(),
            data.get("corrective_action", "").strip(),
            data.get("trigger", "").strip(),
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
    current = get_mitigation(mitigation_id) or {}
    risk_id = data.get("risk_id")
    risk = get_risk(int(risk_id)) if risk_id else None
    execute(
        """
        UPDATE mitigations
        SET risk_id = ?, risk_name = ?, preventive_action = ?, corrective_action = ?,
            trigger = ?, owner = ?, start_date = ?, end_date = ?, resources = ?, status = ?, evidence = ?,
            strategy = ?, progress = ?
        WHERE id = ?
        """,
        (
            int(risk_id) if risk_id else None,
            risk["name"] if risk else data.get("risk_name", "").strip(),
            data.get("preventive_action", current.get("preventive_action", "")).strip(),
            data.get("corrective_action", current.get("corrective_action", "")).strip(),
            data.get("trigger", current.get("trigger", "")).strip(),
            data.get("owner", current.get("owner", "")).strip(),
            iso_date(data.get("start_date", current.get("start_date"))),
            iso_date(data.get("end_date", current.get("end_date"))),
            data.get("resources", current.get("resources", "")).strip(),
            data.get("status", current.get("status", "Planificado")),
            data.get("evidence", current.get("evidence", "")).strip(),
            data.get("strategy", current.get("strategy", "Mitigar")),
            int(data.get("progress") if data.get("progress") is not None else current.get("progress", 0) or 0),
            mitigation_id,
        ),
    )


def delete_mitigation(mitigation_id: int) -> None:
    execute("DELETE FROM mitigations WHERE id = ?", (mitigation_id,))


def calculate_level(probability: float, impact: float) -> str:
    score = probability * impact
    if score <= 4:
        return "Bajo"
    if score <= 9:
        return "Medio"
    if score <= 15:
        return "Alto"
    return "Crítico"


def _normalize_level(level: str | None) -> str:
    raw = (level or "").strip().lower()
    if raw in {"critico", "crítico"}:
        return "Crítico"
    if raw == "alto":
        return "Alto"
    if raw == "medio":
        return "Medio"
    return "Bajo"


def report_summary(project_id: int | None) -> dict[str, Any]:
    if not project_id:
        return {"total": 0, "low": 0, "medium": 0, "high": 0, "critical": 0, "mitigated": 0, "progress": 0}
    rows = query_all("SELECT level, status FROM risks WHERE project_id = ?", (project_id,))
    counts = {"Bajo": 0, "Medio": 0, "Alto": 0, "Crítico": 0}
    mitigated = 0
    for row in rows:
        level = _normalize_level(row["level"])
        counts[level] = counts.get(level, 0) + 1
        if (row["status"] or "").lower() in {"mitigado", "cerrado"}:
            mitigated += 1
    total = len(rows)
    progress = round((mitigated / total) * 100) if total else 0
    return {
        "total": total,
        "low": counts["Bajo"],
        "medium": counts["Medio"],
        "high": counts["Alto"],
        "critical": counts["Crítico"],
        "mitigated": mitigated,
        "progress": progress,
    }
