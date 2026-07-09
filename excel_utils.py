from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any
import unicodedata

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


TEMPLATE_PATH = Path("static/files/template_riesgos_fixed.xlsx")


def _normalize_date(value: Any) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        patterns = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d/%m/%y",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%Y/%m/%d",
        ]
        for pattern in patterns:
            try:
                return datetime.strptime(raw, pattern).date().isoformat()
            except ValueError:
                continue
        return None
    return str(value)


def _normalize_date_with_status(value: Any) -> tuple[str | None, bool]:
    normalized = _normalize_date(value)
    if value in (None, ""):
        return None, True
    if normalized is None:
        return None, False
    if isinstance(value, str):
        raw = value.strip()
        patterns = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d/%m/%y",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%Y/%m/%d",
        ]
        for pattern in patterns:
            try:
                parsed = datetime.strptime(raw, pattern).date().isoformat()
                return parsed, True
            except ValueError:
                continue
        return None, False
    return normalized, True


def _cell_value(cell, wb):
    value = cell.value
    if value is None:
        return None
    if getattr(cell, "is_date", False):
        return value
    if isinstance(value, (int, float)) and cell.number_format and "yy" in str(cell.number_format).lower():
        try:
            return from_excel(value, wb.epoch)
        except Exception:
            return value
    return value


def _sheet_rows(ws, wb, start_row: int = 2) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in ws.iter_rows(min_row=start_row, values_only=False):
        values = [_cell_value(cell, wb) for cell in row]
        if any(value not in (None, "") for value in values):
            rows.append(values)
    return rows


def _normalize_header(value: Any) -> str:
    if value in (None, ""):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(char for char in text if not unicodedata.combining(char))


def _row_map(headers: list[Any], row: list[Any]) -> dict[str, Any]:
    mapped: dict[str, Any] = {}
    for index, header in enumerate(headers):
        key = _normalize_header(header)
        if key:
            mapped[key] = row[index] if index < len(row) else None
    return mapped


def _pick(mapped: dict[str, Any], *aliases: str, default: Any = None) -> Any:
    for alias in aliases:
        key = _normalize_header(alias)
        if key in mapped and mapped[key] not in (None, ""):
            return mapped[key]
    return default


@dataclass
class ImportedWorkbook:
    project: dict[str, Any] | None
    assets: list[dict[str, Any]]
    risks: list[dict[str, Any]]
    mitigations: list[dict[str, Any]]
    warnings: list[str]


def load_template_bytes() -> bytes:
    return TEMPLATE_PATH.read_bytes()


def parse_imported_workbook(file_bytes: bytes) -> ImportedWorkbook:
    wb = load_workbook(BytesIO(file_bytes), data_only=True)

    project_data = None
    warnings: list[str] = []
    if "Proyecto" in wb.sheetnames:
        ws = wb["Proyecto"]
        values = _sheet_rows(ws, wb, start_row=2)
        headers = values[0] if values else []
        if len(values) >= 2:
            row = values[1]
            mapped = _row_map(headers, row)
            raw_start = _pick(mapped, "Fecha Inicio", "Inicio")
            raw_end = _pick(mapped, "Fecha Fin", "Fin")
            start_date, start_ok = _normalize_date_with_status(raw_start)
            end_date, end_ok = _normalize_date_with_status(raw_end)
            if not start_ok and raw_start not in (None, ""):
                warnings.append("Fecha Inicio del proyecto omitida por formato invalido.")
            if not end_ok and raw_end not in (None, ""):
                warnings.append("Fecha Fin del proyecto omitida por formato invalido.")
            project_data = {
                "name": _pick(mapped, "Nombre del proyecto", "Proyecto", "Nombre") or "Proyecto sin nombre",
                "project_type": _pick(mapped, "Tipo", "Tipo de proyecto") or "",
                "company": _pick(mapped, "Empresa", "Compañia", "Compañía") or "",
                "start_date": start_date,
                "end_date": end_date,
                "status": _pick(mapped, "Estado", "Status") or "Activo",
            }

    assets: list[dict[str, Any]] = []
    if "Activos" in wb.sheetnames:
        ws = wb["Activos"]
        rows = _sheet_rows(ws, wb, start_row=2)
        headers = rows[0] if rows else []
        for row in rows[1:]:
            item = _row_map(headers, row)
            assets.append(
                {
                    "name": _pick(item, "Nombre del activo", "Nombre", "Activo") or "",
                    "type": _pick(item, "Tipo") or "",
                    "owner": _pick(item, "Responsable") or "",
                    "value": _pick(item, "Valor del activo", "Valor") or 0,
                    "status": _pick(item, "Estado", "Status") or "Activo",
                }
            )

    risks: list[dict[str, Any]] = []
    if "Riesgos" in wb.sheetnames:
        ws = wb["Riesgos"]
        rows = _sheet_rows(ws, wb, start_row=2)
        headers = rows[0] if rows else []
        for row in rows[1:]:
            item = _row_map(headers, row)
            risks.append(
                {
                    "name": _pick(item, "Riesgo") or "",
                    "asset_name": _pick(item, "Activo") or "",
                    "description": _pick(item, "Descripcion", "Descripción") or "",
                    "cause": _pick(item, "Causa") or "",
                    "consequence": _pick(item, "Consecuencia") or "",
                    "probability": _pick(item, "Probabilidad") or 0,
                    "impact": _pick(item, "Impacto") or 0,
                    "level": _pick(item, "Nivel") or "",
                    "horizon": _pick(item, "Horizonte") or "",
                    "owner": _pick(item, "Responsable") or "",
                    "status": _pick(item, "Estado", "Status") or "Identificado",
                    "strategy": _pick(item, "Estrategia") or "Mitigar",
                }
            )

    mitigations: list[dict[str, Any]] = []
    if "Mitigacion" in wb.sheetnames:
        ws = wb["Mitigacion"]
        rows = _sheet_rows(ws, wb, start_row=2)
        headers = rows[0] if rows else []
        for row in rows[1:]:
            item = _row_map(headers, row)
            start_date, start_ok = _normalize_date_with_status(_pick(item, "Inicio"))
            end_date, end_ok = _normalize_date_with_status(_pick(item, "Fin"))
            if not start_ok and _pick(item, "Inicio") not in (None, ""):
                warnings.append(f"Mitigacion '{_pick(item, 'Riesgo') or ''}' con Inicio invalido omitido.")
            if not end_ok and _pick(item, "Fin") not in (None, ""):
                warnings.append(f"Mitigacion '{_pick(item, 'Riesgo') or ''}' con Fin invalido omitido.")
            mitigations.append(
                {
                    "risk_name": _pick(item, "Riesgo") or "",
                    "preventive_action": _pick(item, "Accion preventiva", "Accion Preventiva") or "",
                    "corrective_action": _pick(item, "Accion correctiva", "Accion Correctiva") or "",
                    "owner": _pick(item, "Responsable") or "",
                    "start_date": start_date,
                    "end_date": end_date,
                    "resources": _pick(item, "Recursos") or "",
                    "status": _pick(item, "Estado", "Status") or "Planificado",
                    "evidence": _pick(item, "Evidencia") or "",
                    "strategy": _pick(item, "Estrategia") or "Mitigar",
                    "progress": _pick(item, "Avance") or 0,
                }
            )

    return ImportedWorkbook(project=project_data, assets=assets, risks=risks, mitigations=mitigations, warnings=warnings)
