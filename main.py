from __future__ import annotations

import os
import json
import subprocess
import tempfile
import unicodedata
from io import BytesIO
from xml.sax.saxutils import escape as xml_escape
from pathlib import Path
from datetime import date, timedelta
from typing import Any

from flask import (
    Flask,
    abort,
    flash,
    get_flashed_messages,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image as ReportImage
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from db import (
    close_db,
    create_asset,
    create_mitigation,
    create_project,
    create_project_deliverable,
    create_project_safeguard,
    create_project_schedule,
    create_project_threat,
    create_project_role,
    create_risk,
    delete_project_role,
    delete_project_safeguard,
    delete_project_threat,
    delete_asset,
    delete_mitigation,
    delete_project,
    delete_project_deliverable,
    delete_project_schedule,
    delete_risk,
    get_asset,
    get_mitigation,
    get_project,
    get_project_deliverable,
    get_project_safeguard,
    get_project_schedule,
    get_project_threat,
    get_project_role,
    get_risk,
    init_db,
    list_assets,
    list_assets_for_select,
    list_mitigations,
    list_projects,
    list_project_deliverables,
    list_project_safeguards,
    list_project_schedule,
    list_project_threats,
    list_risk_names,
    list_project_roles,
    list_risks,
    report_summary,
    reset_db,
    update_asset,
    update_mitigation,
    update_project,
    update_project_deliverable,
    update_project_safeguard,
    update_project_schedule,
    update_project_threat,
    update_project_role,
    update_risk,
)
from excel_utils import load_template_bytes, parse_imported_workbook


BASE_DIR = Path(__file__).resolve().parent


def _lookup_key(value: Any) -> str:
    if value in (None, ""):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(char for char in text if not unicodedata.combining(char) and char.isalnum())


def _safe_text(value: Any, default: str = "-") -> str:
    if value in (None, ""):
        return default
    return str(value)


def _build_cronograma_phases() -> list[dict[str, Any]]:
    return [
        {
            "phase": "Inicio y planificación",
            "risks": ["R01", "R02", "R03", "R04", "R05", "R22"],
            "control": "Aprobar Project Charter, definir alcance, elaborar cronograma, matriz de recursos y control de cambios.",
            "start_period": 1,
            "end_period": 2,
        },
        {
            "phase": "Levantamiento y comprensión del negocio",
            "risks": ["R06", "R07", "R08"],
            "control": "Sesiones con Product Owner, validación de procesos contables y criterios documentales.",
            "start_period": 3,
            "end_period": 3,
        },
        {
            "phase": "Datos, OCR y Data Mining",
            "risks": ["R09", "R10", "R12", "R13", "R14"],
            "control": "Recolección mínima por categoría, limpieza, etiquetado, pruebas OCR y validación de dataset.",
            "start_period": 4,
            "end_period": 5,
        },
        {
            "phase": "Modelado IA",
            "risks": ["R15", "R16", "R17"],
            "control": "Comparación de modelos, evaluación con métricas ML, ajuste y reentrenamiento.",
            "start_period": 6,
            "end_period": 7,
        },
        {
            "phase": "Desarrollo web y base de datos",
            "risks": ["R18", "R19", "R20", "R21", "R23"],
            "control": "Contratos API, revisión de arquitectura, controles de acceso, pruebas unitarias e integración.",
            "start_period": 8,
            "end_period": 9,
        },
        {
            "phase": "Pruebas y validación",
            "risks": ["R24", "R25"],
            "control": "Plan QA, pruebas de API, pruebas funcionales, UAT y actas de aceptación.",
            "start_period": 10,
            "end_period": 10,
        },
        {
            "phase": "Despliegue, capacitación y cierre",
            "risks": ["R26", "R27", "R28", "R29", "R30"],
            "control": "Despliegue controlado, monitoreo, manuales, capacitación, transferencia y cierre formal.",
            "start_period": 11,
            "end_period": 12,
        },
    ]


def _build_project_report_context(project_id: int) -> dict[str, Any]:
    project = get_project(project_id)
    risks = list_risks(project_id)
    matrix_rows = []
    for risk in risks:
        score = float(risk.get("probability") or 0) * float(risk.get("impact") or 0)
        level = _safe_text(risk.get("level"))
        priority = "Crítica" if level.lower() in {"critico", "crítico"} else "Alta" if level == "Alto" else "Moderada" if level == "Medio" else "Baja"
        matrix_rows.append(
            {
                "risk": risk,
                "score": score,
                "priority": priority,
            }
        )

    return {
        "project": project,
        "roles": list_project_roles(project_id),
        "deliverables": list_project_deliverables(project_id),
        "threats": list_project_threats(project_id),
        "safeguards": list_project_safeguards(project_id),
        "assets": list_assets(project_id),
        "risks": risks,
        "matrix_rows": matrix_rows,
        "mitigations": list_mitigations(project_id),
        "summary": report_summary(project_id),
        "cronograma_phases": _build_cronograma_phases(),
    }


def _level_fill(level: str | None) -> colors.Color:
    normalized = (level or "").strip().lower()
    if normalized in {"critico", "crítico"}:
        return colors.HexColor("#ef4444")
    if normalized == "alto":
        return colors.HexColor("#f97316")
    if normalized == "medio":
        return colors.HexColor("#f59e0b")
    return colors.HexColor("#22c55e")


def _build_project_pdf(context: dict[str, Any]) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Proyecto {context['project']['code']} - {context['project']['name']}",
        author="Sistema de Gestión de Riesgos",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ProjectTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#1f3b73"),
        spaceAfter=8,
    )
    section_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=14,
        textColor=colors.HexColor("#1f3b73"),
        spaceBefore=8,
        spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "BodySmall",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
    )
    small_style = ParagraphStyle(
        "Small",
        parent=body_style,
        fontSize=7.5,
        leading=9,
    )

    def p(value: Any, style: ParagraphStyle = body_style) -> Paragraph:
        return Paragraph(xml_escape(_safe_text(value)), style)

    story: list[Any] = []
    project = context["project"] or {}
    story.append(Paragraph(f"Proyecto {xml_escape(_safe_text(project.get('code')))}", title_style))
    story.append(Paragraph(xml_escape(_safe_text(project.get("name"))), styles["Heading1"]))
    story.append(Spacer(1, 4))

    summary_data = [
        [p("Código"), p(project.get("code"))],
        [p("Tipo"), p(project.get("project_type"))],
        [p("Empresa"), p(project.get("company"))],
        [p("Inicio"), p(project.get("start_date"))],
        [p("Fin"), p(project.get("end_date"))],
        [p("Estado"), p(project.get("status"))],
        [p("Riesgos"), p(context["summary"]["total"])],
        [p("Mitigados"), p(context["summary"]["mitigated"])],
    ]
    summary_table = Table(summary_data, colWidths=[34 * mm, 75 * mm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbeafe")),
                ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#eff6ff")),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("LEADING", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(summary_table)
    story.append(Spacer(1, 6))

    sections = [
        ("Entregables", [["Entregable"], *[[p(item.get("deliverable"))] for item in context["deliverables"]]]),
        ("Roles", [["Rol", "Adquisición", "Participación en riesgos"], *[[p(item.get("role")), p(item.get("acquisition_type")), p(item.get("risk_participation"))] for item in context["roles"]]]),
        ("Activos", [["Activo", "Tipo de activo", "Valor"], *[[p(item.get("name")), p(item.get("type")), p(item.get("value"))] for item in context["assets"]]]),
        ("Amenazas", [["Código", "Amenaza", "Activo afectado", "Ejemplo"], *[[p(item.get("code")), p(item.get("threat")), p(item.get("affected_asset")), p(item.get("example"))] for item in context["threats"]]]),
        ("Salvaguardas", [["Amenaza", "Salvaguarda"], *[[p(item.get("threat_name") or item.get("threat_code")), p(item.get("safeguard"))] for item in context["safeguards"]]]),
        ("Riesgos", [["Código", "Riesgo", "Activo", "P", "I", "Nivel", "Estado"], *[[p(item.get("code")), p(item.get("name")), p(item.get("asset_name")), p(item.get("probability")), p(item.get("impact")), p(item.get("level")), p(item.get("status"))] for item in context["risks"]]]),
    ]

    def build_section_table(title: str, data: list[list[Paragraph]]) -> None:
        story.append(Paragraph(xml_escape(title), section_style))
        if len(data) == 1:
            story.append(Paragraph("Sin registros.", small_style))
            story.append(Spacer(1, 4))
            return
        table = Table(data, repeatRows=1)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("LEADING", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        if title == "Riesgos":
            style_cmds.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1d4ed8")))
            for idx, row in enumerate(context["matrix_rows"], start=1):
                style_cmds.append(("BACKGROUND", (0, idx), (-1, idx), _level_fill(row["risk"].get("level"))))
                style_cmds.append(("TEXTCOLOR", (0, idx), (-1, idx), colors.white if _level_fill(row["risk"].get("level")) != colors.HexColor("#f59e0b") else colors.black))
        table.setStyle(TableStyle(style_cmds))
        story.append(table)
        story.append(Spacer(1, 6))

    for title, data in sections:
        build_section_table(title, data)

    matrix_block: list[Any] = [Paragraph("Matriz de riesgos", section_style)]
    matrix_data = [[p("Riesgo"), p("Activo afectado"), p("P"), p("I"), p("Nivel"), p("Prioridad")]]
    for row in context["matrix_rows"]:
        risk = row["risk"]
        matrix_data.append([
            p(risk.get("name")),
            p(risk.get("asset_name") or risk.get("asset_label")),
            p(risk.get("probability")),
            p(risk.get("impact")),
            p(risk.get("level")),
            p(row["priority"]),
        ])
    matrix_table = Table(matrix_data, repeatRows=1, colWidths=[58 * mm, 60 * mm, 18 * mm, 18 * mm, 28 * mm, 28 * mm])
    matrix_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f3b6e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]
    for idx, row in enumerate(context["matrix_rows"], start=1):
        fill = _level_fill(row["risk"].get("level"))
        matrix_style.append(("BACKGROUND", (4, idx), (5, idx), fill))
        matrix_style.append(("TEXTCOLOR", (4, idx), (5, idx), colors.white))
    matrix_table.setStyle(TableStyle(matrix_style))
    story.append(matrix_table)
    story.append(Spacer(1, 8))

    story.append(Paragraph("Cronograma visual", section_style))
    timeline_data = [[p(phase["phase"], small_style) for phase in context["cronograma_phases"]]]
    timeline_table = Table(timeline_data, colWidths=[34 * mm] * len(context["cronograma_phases"]))
    timeline_style = [
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]
    timeline_table.setStyle(TableStyle(timeline_style))
    story.append(timeline_table)

    risk_row = [[p(", ".join(phase["risks"]), small_style) for phase in context["cronograma_phases"]]]
    risk_table = Table(risk_row, colWidths=[34 * mm] * len(context["cronograma_phases"]))
    risk_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eff6ff")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(risk_table)

    control_row = [[p(phase["control"], small_style) for phase in context["cronograma_phases"]]]
    control_table = Table(control_row, colWidths=[34 * mm] * len(context["cronograma_phases"]))
    control_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(control_table)
    story.append(Spacer(1, 6))

    story.append(Paragraph("Mitigaciones", section_style))
    mitigation_data = [[p("Riesgo"), p("Acción preventiva"), p("Acción correctiva / contingencia"), p("Disparador")]]
    for item in context["mitigations"]:
        mitigation_data.append([
            p(item.get("risk_name") or item.get("risk_label")),
            p(item.get("preventive_action")),
            p(item.get("corrective_action")),
            p(item.get("trigger")),
        ])
    mitigation_table = Table(mitigation_data, repeatRows=1, colWidths=[35 * mm, 70 * mm, 70 * mm, 55 * mm])
    mitigation_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.25),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]))
    story.append(mitigation_table)

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def _project_cronograma_phases() -> list[dict[str, Any]]:
    return [
        {
            "phase": "Inicio y planificación",
            "risks": ["R01", "R02", "R03", "R04", "R05", "R22"],
            "control": "Aprobar Project Charter, definir alcance, elaborar cronograma, matriz de recursos y control de cambios.",
            "start_period": 1,
            "end_period": 2,
        },
        {
            "phase": "Levantamiento y comprensión del negocio",
            "risks": ["R06", "R07", "R08"],
            "control": "Sesiones con Product Owner, validación de procesos contables y criterios documentales.",
            "start_period": 3,
            "end_period": 3,
        },
        {
            "phase": "Datos, OCR y Data Mining",
            "risks": ["R09", "R10", "R12", "R13", "R14"],
            "control": "Recolección mínima por categoría, limpieza, etiquetado, pruebas OCR y validación de dataset.",
            "start_period": 4,
            "end_period": 5,
        },
        {
            "phase": "Modelado IA",
            "risks": ["R15", "R16", "R17"],
            "control": "Comparación de modelos, evaluación con métricas ML, ajuste y reentrenamiento.",
            "start_period": 6,
            "end_period": 7,
        },
        {
            "phase": "Desarrollo web y base de datos",
            "risks": ["R18", "R19", "R20", "R21", "R23"],
            "control": "Contratos API, revisión de arquitectura, controles de acceso, pruebas unitarias e integración.",
            "start_period": 8,
            "end_period": 9,
        },
        {
            "phase": "Pruebas y validación",
            "risks": ["R24", "R25"],
            "control": "Plan QA, pruebas de API, pruebas funcionales, UAT y actas de aceptación.",
            "start_period": 10,
            "end_period": 10,
        },
        {
            "phase": "Despliegue, capacitación y cierre",
            "risks": ["R26", "R27", "R28", "R29", "R30"],
            "control": "Despliegue controlado, monitoreo, manuales, capacitación, transferencia y cierre formal.",
            "start_period": 11,
            "end_period": 12,
        },
    ]


def _project_report_context(project_id: int) -> dict[str, Any]:
    project = get_project(project_id)
    risks = list_risks(project_id)
    matrix_rows = []
    for risk in risks:
        score = float(risk.get("probability") or 0) * float(risk.get("impact") or 0)
        level = (risk.get("level") or "").strip()
        normalized = level.lower()
        priority = "Crítica" if normalized in {"critico", "crítico"} else "Alta" if normalized in {"alto", "alta"} else "Moderada" if normalized == "medio" else "Baja"
        matrix_rows.append({"risk": risk, "score": score, "priority": priority})

    return {
        "project": project,
        "roles": list_project_roles(project_id),
        "deliverables": list_project_deliverables(project_id),
        "threats": list_project_threats(project_id),
        "safeguards": list_project_safeguards(project_id),
        "assets": list_assets(project_id),
        "risks": risks,
        "matrix_rows": matrix_rows,
        "mitigations": list_mitigations(project_id),
        "summary": report_summary(project_id),
        "cronograma_phases": list_project_schedule(project_id),
    }


def _project_level_fill(level: str | None) -> colors.Color:
    normalized = (level or "").strip().lower()
    if normalized in {"critico", "crítico"}:
        return colors.HexColor("#ef4444")
    if normalized in {"alto", "alta"}:
        return colors.HexColor("#f97316")
    if normalized == "medio":
        return colors.HexColor("#eab308")
    return colors.HexColor("#00b050")


def _browser_executable() -> str | None:
    candidates = [
        Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def _heatmap_chart_image(context: dict[str, Any], temp_files: list[Path]) -> Path | None:
    browser = _browser_executable()
    if not browser:
        return None

    points = []
    for risk in context["risks"]:
        level_key = _lookup_key(risk.get("level") or "Bajo")
        color = "#00b050" if level_key == "bajo" else "#eab308" if level_key == "medio" else "#f97316" if level_key in {"alto", "alta"} else "#ef4444"
        points.append(
            {
                "x": float(risk.get("probability") or 0),
                "y": float(risk.get("impact") or 0),
                "label": _safe_text(risk.get("name")),
                "color": color,
            }
        )

    html = f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
  <style>
    html, body {{ margin: 0; width: 1100px; height: 650px; background: #ffffff; font-family: Arial, sans-serif; }}
    .frame {{ width: 1040px; height: 590px; padding: 28px 30px; box-sizing: border-box; }}
    .title {{ color: #1f3b73; font-size: 22px; font-weight: 700; margin: 0 0 16px; }}
    .chart-shell {{ width: 980px; height: 520px; border: 1px solid #d8dee9; border-radius: 8px; padding: 16px; box-sizing: border-box; background: #ffffff; }}
    canvas {{ width: 100% !important; height: 100% !important; }}
  </style>
</head>
<body>
  <div class="frame">
    <div class="title">Mapa de calor visual</div>
    <div class="chart-shell"><canvas id="heatChart"></canvas></div>
  </div>
  <script>
    const points = {json.dumps(points, ensure_ascii=False)};
    const diagonalHeatBackground = {{
      id: "diagonalHeatBackground",
      beforeDraw(chart) {{
        const {{ ctx, chartArea }} = chart;
        if (!chartArea) return;
        const {{ left, right, top, bottom, width, height }} = chartArea;
        const gradient = ctx.createLinearGradient(left, bottom, right, top);
        gradient.addColorStop(0.0, "#00b050");
        gradient.addColorStop(0.35, "#eab308");
        gradient.addColorStop(0.70, "#f97316");
        gradient.addColorStop(1.0, "#ef4444");
        ctx.save();
        ctx.fillStyle = gradient;
        ctx.fillRect(left, top, width, height);
        ctx.restore();
      }},
    }};
    new Chart(document.getElementById("heatChart"), {{
      type: "scatter",
      data: {{
        datasets: [{{
          label: "Riesgos",
          data: points,
          pointRadius: 8,
          pointHoverRadius: 10,
          backgroundColor: points.map((point) => point.color),
          borderColor: "#0b0f19",
          borderWidth: 2,
        }}]
      }},
      plugins: [diagonalHeatBackground],
      options: {{
        animation: false,
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{
            callbacks: {{
              label: (context) => `${{context.raw.label}} (P: ${{context.raw.x}}, I: ${{context.raw.y}})`,
            }},
          }},
        }},
        scales: {{
          x: {{
            min: 0.5,
            max: 5.5,
            ticks: {{ stepSize: 1 }},
            grid: {{ color: "rgba(255,255,255,0.12)" }},
            title: {{ display: true, text: "Probabilidad", color: "#9ca3af", font: {{ weight: "bold", size: 12 }} }},
          }},
          y: {{
            min: 0.5,
            max: 5.5,
            ticks: {{ stepSize: 1 }},
            grid: {{ color: "rgba(255,255,255,0.12)" }},
            title: {{ display: true, text: "Impacto", color: "#9ca3af", font: {{ weight: "bold", size: 12 }} }},
          }},
        }},
      }}
    }});
  </script>
</body>
</html>"""

    html_file = Path(tempfile.NamedTemporaryFile(delete=False, suffix=".html").name)
    png_file = Path(tempfile.NamedTemporaryFile(delete=False, suffix=".png").name)
    html_file.write_text(html, encoding="utf-8")
    temp_files.extend([html_file, png_file])
    command = [
        browser,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--disable-extensions",
        "--hide-scrollbars",
        "--window-size=1100,650",
        "--virtual-time-budget=2500",
        f"--screenshot={png_file}",
        html_file.as_uri(),
    ]
    try:
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
    except (subprocess.SubprocessError, OSError):
        return None
    return png_file if png_file.exists() and png_file.stat().st_size > 0 else None


def _build_project_pdf_clean(context: dict[str, Any]) -> bytes:
    temp_files: list[Path] = []
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Proyecto {context['project']['code']} - {context['project']['name']}",
        author="Sistema de Gestión de Riesgos",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("ProjectTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=colors.HexColor("#1f3b73"), spaceAfter=8)
    section_style = ParagraphStyle("SectionTitle", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=colors.HexColor("#1f3b73"), spaceBefore=8, spaceAfter=6)
    body_style = ParagraphStyle("BodySmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=10)
    small_style = ParagraphStyle("Small", parent=body_style, fontSize=7.5, leading=9)

    def p(value: Any, style: ParagraphStyle = body_style) -> Paragraph:
        return Paragraph(xml_escape(_safe_text(value)), style)

    project = context["project"] or {}
    story: list[Any] = [
        Paragraph(f"Proyecto {xml_escape(_safe_text(project.get('code')))}", title_style),
        Paragraph(xml_escape(_safe_text(project.get("name"))), styles["Heading1"]),
        Spacer(1, 4),
    ]

    summary_data = [
        [p("Código"), p(project.get("code"))],
        [p("Tipo"), p(project.get("project_type"))],
        [p("Empresa"), p(project.get("company"))],
        [p("Inicio"), p(project.get("start_date"))],
        [p("Fin"), p(project.get("end_date"))],
        [p("Estado"), p(project.get("status"))],
        [p("Riesgos"), p(context["summary"]["total"])],
        [p("Mitigados"), p(context["summary"]["mitigated"])],
    ]
    summary_table = Table(summary_data, colWidths=[34 * mm, 75 * mm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbeafe")),
        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#eff6ff")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("LEADING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 6))

    def add_table(title: str, headers: list[str], rows: list[list[Any]], widths: list[float] | None = None) -> None:
        story.append(Paragraph(title, section_style))
        if not rows:
            story.append(Paragraph("Sin registros.", small_style))
            story.append(Spacer(1, 4))
            return
        table_data = [headers] + [[p(value) for value in row] for row in rows]
        table = Table(table_data, repeatRows=1, colWidths=widths)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.25),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ]
        table.setStyle(TableStyle(style))
        story.append(table)
        story.append(Spacer(1, 5))

    def heat_color(probability: int, impact: int) -> colors.Color:
        score = probability * impact
        if score >= 16:
            return colors.HexColor("#ef4444")
        if score >= 11:
            return colors.HexColor("#f97316")
        if score >= 6:
            return colors.HexColor("#eab308")
        return colors.HexColor("#00b050")

    def add_heatmap_visual() -> None:
        chart_image = _heatmap_chart_image(context, temp_files)
        if chart_image:
            block: list[Any] = []
            block.append(ReportImage(str(chart_image), width=210 * mm, height=124 * mm))
            block.append(Spacer(1, 8))
            story.append(KeepTogether(block))
            return

        block: list[Any] = [Paragraph("Mapa de calor visual", section_style)]
        risks_by_cell: dict[tuple[int, int], list[str]] = {}
        for risk in context["risks"]:
            probability = int(float(risk.get("probability") or 0))
            impact = int(float(risk.get("impact") or 0))
            if 1 <= probability <= 5 and 1 <= impact <= 5:
                label = _safe_text(risk.get("code") or risk.get("name"))
                risks_by_cell.setdefault((probability, impact), []).append(label)

        axis_style = ParagraphStyle("HeatAxis", parent=small_style, fontName="Helvetica-Bold", fontSize=7.4, leading=8.5, textColor=colors.white, alignment=1)
        cell_style = ParagraphStyle("HeatCell", parent=small_style, fontName="Helvetica-Bold", fontSize=7.4, leading=8.5, textColor=colors.black, alignment=1)
        empty_style = ParagraphStyle("HeatEmpty", parent=small_style, fontSize=7.4, leading=8.5, textColor=colors.transparent, alignment=1)

        legend = Table(
            [[
                p("Bajo", small_style),
                p("Medio", small_style),
                p("Alto", small_style),
                p("Crítico", small_style),
            ]],
            colWidths=[24 * mm, 24 * mm, 24 * mm, 24 * mm],
        )
        legend.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#00b050")),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#eab308")),
            ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#f97316")),
            ("BACKGROUND", (3, 0), (3, 0), colors.HexColor("#ef4444")),
            ("TEXTCOLOR", (0, 0), (0, 0), colors.black),
            ("TEXTCOLOR", (1, 0), (1, 0), colors.black),
            ("TEXTCOLOR", (2, 0), (3, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ]))
        block.append(legend)
        block.append(Spacer(1, 4))

        data: list[list[Any]] = [[Paragraph("Impacto \\ Probabilidad", axis_style), *[Paragraph(str(i), axis_style) for i in range(1, 6)]]]
        for impact in range(5, 0, -1):
            row: list[Any] = [Paragraph(str(impact), axis_style)]
            for probability in range(1, 6):
                labels = risks_by_cell.get((probability, impact), [])
                row.append(Paragraph("<br/>".join(labels[:4]), cell_style) if labels else Paragraph("", empty_style))
            data.append(row)

        table = Table(data, colWidths=[40 * mm, 30 * mm, 30 * mm, 30 * mm, 30 * mm, 30 * mm], rowHeights=[11 * mm] + [18 * mm] * 5)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#111827")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.white),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#111827")),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
        for row_idx, impact in enumerate(range(5, 0, -1), start=1):
            for col_idx, probability in enumerate(range(1, 6), start=1):
                fill = heat_color(probability, impact)
                style.append(("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), fill))
        table.setStyle(TableStyle(style))
        block.append(table)
        block.append(Spacer(1, 8))
        story.append(KeepTogether(block))

    def add_cronograma_visual() -> None:
        story.append(Paragraph("Cronograma visual del proyecto", section_style))
        phases = context["cronograma_phases"]
        data: list[list[Any]] = [[p("Actividades", small_style), *[p(period, small_style) for period in range(1, 13)]]]
        for phase in phases:
            risks_text = ", ".join(phase.get("risks_list") or []) or _safe_text(phase.get("risks"))
            activity = [
                Paragraph(
                    f"<b>{xml_escape(_safe_text(phase.get('phase')))}</b><br/>{xml_escape(risks_text)}<br/><font color='#64748b'>{xml_escape(_safe_text(phase.get('control')))}</font>",
                    small_style,
                )
            ]
            start = int(phase.get("start_period") or 0)
            end = int(phase.get("end_period") or start)
            for period in range(1, 13):
                activity.append(p("●" if start <= period <= end else "", small_style))
            data.append(activity)

        col_widths = [78 * mm] + [13.5 * mm] * 12
        table = Table(data, colWidths=col_widths, repeatRows=1)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0a0df")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
            ("FONTSIZE", (0, 0), (-1, -1), 6.6),
            ("LEFTPADDING", (0, 1), (0, -1), 5),
            ("RIGHTPADDING", (0, 1), (0, -1), 5),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ]
        for row_idx, phase in enumerate(phases, start=1):
            start = int(phase.get("start_period") or 0)
            end = int(phase.get("end_period") or start)
            for period in range(1, 13):
                if start <= period <= end:
                    style.append(("BACKGROUND", (period, row_idx), (period, row_idx), colors.HexColor("#f5e86d")))
                    style.append(("TEXTCOLOR", (period, row_idx), (period, row_idx), colors.HexColor("#f5e86d")))
        table.setStyle(TableStyle(style))
        story.append(table)
        story.append(Spacer(1, 8))

    add_table(
        "Entregables",
        ["Entregable", "Descripción", "Responsable principal"],
        [[item.get("deliverable"), item.get("description"), item.get("main_responsible")] for item in context["deliverables"]],
        [58 * mm, 110 * mm, 62 * mm],
    )
    add_table("Roles", ["Rol", "Adquisición", "Participación en riesgos"], [[item.get("role"), item.get("acquisition_type"), item.get("risk_participation")] for item in context["roles"]], [62 * mm, 38 * mm, 95 * mm])
    add_table("Activos", ["Activo", "Tipo de activo", "Valor"], [[item.get("name"), item.get("type"), item.get("value")] for item in context["assets"]], [78 * mm, 65 * mm, 52 * mm])
    add_table("Amenazas", ["Código", "Amenaza", "Activo afectado", "Ejemplo"], [[item.get("code"), item.get("threat"), item.get("affected_asset"), item.get("example")] for item in context["threats"]], [24 * mm, 45 * mm, 85 * mm, 40 * mm])
    add_table("Salvaguardas", ["Amenaza", "Salvaguarda"], [[item.get("threat_name") or item.get("threat_code"), item.get("safeguard")] for item in context["safeguards"]], [55 * mm, 140 * mm])
    add_table("Riesgos", ["Código", "Riesgo", "Activo", "P", "I", "Nivel", "Estado"], [[item.get("code"), item.get("name"), item.get("asset_name"), item.get("probability"), item.get("impact"), item.get("level"), item.get("status")] for item in context["risks"]], [22 * mm, 68 * mm, 52 * mm, 12 * mm, 12 * mm, 20 * mm, 22 * mm])

    add_heatmap_visual()

    matrix_block: list[Any] = [Paragraph("Matriz de riesgos", section_style)]
    matrix_data = [[p("Riesgo"), p("Activo afectado"), p("P"), p("I"), p("Nivel"), p("Prioridad")]]
    for row in context["matrix_rows"]:
        risk = row["risk"]
        matrix_data.append([
            p(risk.get("name")),
            p(risk.get("asset_name") or risk.get("asset_label")),
            p(risk.get("probability")),
            p(risk.get("impact")),
            p(risk.get("level")),
            p(row["priority"]),
        ])
    matrix_table = Table(matrix_data, repeatRows=1, colWidths=[58 * mm, 60 * mm, 18 * mm, 18 * mm, 28 * mm, 28 * mm])
    matrix_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f3b6e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]
    for idx, row in enumerate(context["matrix_rows"], start=1):
        fill = _project_level_fill(row["risk"].get("level"))
        matrix_style.append(("BACKGROUND", (4, idx), (5, idx), fill))
        matrix_style.append(("TEXTCOLOR", (4, idx), (5, idx), colors.white))
    matrix_table.setStyle(TableStyle(matrix_style))
    matrix_block.append(matrix_table)
    matrix_block.append(Spacer(1, 8))
    story.append(KeepTogether(matrix_block))

    story.append(PageBreak())
    add_cronograma_visual()

    add_table("Mitigaciones", ["Riesgo", "Acción preventiva", "Acción correctiva / contingencia", "Disparador"], [[item.get("risk_name") or item.get("risk_label"), item.get("preventive_action"), item.get("corrective_action"), item.get("trigger")] for item in context["mitigations"]], [35 * mm, 70 * mm, 70 * mm, 55 * mm])

    try:
        doc.build(story)
        pdf_bytes = buffer.getvalue()
    finally:
        buffer.close()
        for temp_file in temp_files:
            try:
                temp_file.unlink(missing_ok=True)
            except OSError:
                pass
    return pdf_bytes


def create_app() -> Flask:
    app = Flask(__name__, instance_relative_config=True)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "gestion-proyectos-dev")
    app.config["DATABASE"] = os.path.join(app.instance_path, "gestion_proyectos.sqlite3")
    os.makedirs(app.instance_path, exist_ok=True)

    with app.app_context():
        init_db()

    app.teardown_appcontext(close_db)

    def _session_project_id() -> int | None:
        raw_project_id = session.get("project_id")
        try:
            return int(raw_project_id) if raw_project_id is not None else None
        except (TypeError, ValueError):
            return None

    @app.before_request
    def keep_project_selection_valid() -> None:
        page_endpoints = {
            "index",
            "proyectos",
            "select_project",
            "cronograma",
            "activos",
            "entregables",
            "riesgos",
            "roles",
            "amenazas",
            "matriz",
            "mitigacion",
            "pdf_proyecto",
            "reportes",
        }
        if request.endpoint not in page_endpoints:
            return

        projects = list_projects()
        if not projects:
            session.pop("project_id", None)
            if request.method == "GET":
                flash("No hay proyectos registrados. Crea o importa uno para continuar.", "warning")
            return

        selected = _session_project_id()
        if selected is None or not any(project["id"] == selected for project in projects):
            session["project_id"] = projects[0]["id"]

    @app.context_processor
    def inject_globals() -> dict[str, Any]:
        projects = list_projects()
        selected_project_id = _session_project_id()
        current_project = get_project(selected_project_id) if selected_project_id else None
        today_iso = date.today().isoformat()
        default_end_date_iso = (date.today() + timedelta(days=30)).isoformat()
        return {
            "projects": projects,
            "current_project": current_project,
            "selected_project_id": selected_project_id,
            "flash_messages": get_flashed_messages(with_categories=True),
            "today_iso": today_iso,
            "default_end_date_iso": default_end_date_iso,
        }

    def _selected_project_id() -> int | None:
        return _session_project_id()

    def _require_project() -> int | None:
        project_id = _selected_project_id()
        return project_id

    def _redirect_back(default_endpoint: str) -> str:
        next_url = request.form.get("next") or request.args.get("next") or request.referrer
        return next_url or url_for(default_endpoint)

    def _set_selected_project(project_id: int) -> None:
        session["project_id"] = project_id

    def _render_projects(editing_project: dict[str, Any] | None = None):
        return render_template(
            "proyectos.html",
            title="Proyectos",
            active_page="proyectos",
            editing_project=editing_project,
            template_ready=bool((BASE_DIR / "static" / "files" / "template_riesgos.xlsx").exists()),
        )

    def _render_assets(editing_asset: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "activos.html",
            title="Activos",
            active_page="activos",
            current_assets=list_assets(project_id),
            editing_asset=editing_asset,
            asset_options=list_assets_for_select(project_id),
        )

    def _render_deliverables(editing_deliverable: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "entregables.html",
            title="Entregables",
            active_page="entregables",
            project_deliverables=list_project_deliverables(project_id),
            editing_deliverable=editing_deliverable,
        )

    def _render_cronograma():
        project_id = _selected_project_id()
        return render_template(
            "cronograma.html",
            title="Cronograma",
            active_page="cronograma",
            cronograma_phases=list_project_schedule(project_id),
        )

    def _render_risks(editing_risk: dict[str, Any] | None = None, template_name: str = "riesgos.html"):
        project_id = _selected_project_id()
        return render_template(
            template_name,
            title="Riesgos",
            active_page="riesgos",
            current_risks=list_risks(project_id),
            editing_risk=editing_risk,
            asset_options=list_assets_for_select(project_id),
        )

    def _render_roles(editing_role: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "roles.html",
            title="Roles",
            active_page="roles",
            project_roles=list_project_roles(project_id),
            editing_role=editing_role,
        )

    def _render_threats(editing_threat: dict[str, Any] | None = None, editing_safeguard: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "amenazas.html",
            title="Amenazas",
            active_page="amenazas",
            project_threats=list_project_threats(project_id),
            project_safeguards=list_project_safeguards(project_id),
            threat_options=list_project_threats(project_id),
            asset_options=list_assets_for_select(project_id),
            editing_threat=editing_threat,
            editing_safeguard=editing_safeguard,
        )

    def _render_matrix():
        project_id = _selected_project_id()
        risks = list_risks(project_id)
        return render_template(
            "matriz.html",
            title="Matriz de riesgos",
            active_page="matriz",
            matrix_risks=risks,
        )

    def _render_mitigations(editing_mitigation: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "mitigacion.html",
            title="Mitigación",
            active_page="mitigacion",
            current_mitigations=list_mitigations(project_id),
            editing_mitigation=editing_mitigation,
            risk_options=list_risk_names(project_id),
        )

    def _render_reports():
        project_id = _selected_project_id()
        summary = report_summary(project_id)
        risks = list_risks(project_id)
        mitigation_rows = list_mitigations(project_id)
        return render_template(
            "reportes.html",
            title="Reportes",
            active_page="reportes",
            summary=summary,
            report_risks=risks,
            report_mitigations=mitigation_rows,
        )

    def _render_pdf_project():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        context = _project_report_context(project_id)
        return render_template(
            "pdf_proyecto.html",
            title="PDF del proyecto",
            active_page="pdf_proyecto",
            report_context=context,
        )

    @app.route("/")
    def index():
        return redirect(url_for("proyectos"))

    @app.route("/select-project", methods=["POST"])
    def select_project():
        project_id = request.form.get("project_id", type=int)
        if project_id and get_project(project_id):
            _set_selected_project(project_id)
            flash("Proyecto seleccionado.", "success")
        else:
            flash("No se pudo cambiar el proyecto.", "error")
        return redirect(_redirect_back("index"))

    @app.route("/proyectos", methods=["GET", "POST"])
    def proyectos():
        if request.method == "POST":
            project_id = request.form.get("project_id", type=int)
            payload = {
                "name": request.form.get("name", ""),
                "project_type": request.form.get("project_type", ""),
                "company": request.form.get("company", ""),
                "start_date": request.form.get("start_date"),
                "end_date": request.form.get("end_date"),
                "status": request.form.get("status", "Planificado"),
            }

            if project_id:
                if not get_project(project_id):
                    flash("El proyecto no existe.", "error")
                else:
                    update_project(project_id, payload)
                    flash("Proyecto actualizado.", "success")
            else:
                created = create_project(payload)
                _set_selected_project(created["id"])
                flash("Proyecto creado.", "success")
            return redirect(url_for("proyectos"))

        edit_id = request.args.get("edit", type=int)
        editing_project = get_project(edit_id) if edit_id else None
        return _render_projects(editing_project=editing_project)

    @app.route("/proyectos/delete/<int:project_id>", methods=["POST"])
    def delete_project_route(project_id: int):
        if not get_project(project_id):
            flash("El proyecto no existe.", "error")
            return redirect(url_for("proyectos"))

        delete_project(project_id)
        if session.get("project_id") == project_id:
            projects = list_projects()
            if projects:
                session["project_id"] = projects[0]["id"]
            else:
                session.pop("project_id", None)
        flash("Proyecto eliminado.", "success")
        return redirect(url_for("proyectos"))

    @app.route("/proyectos/template")
    def proyectos_template():
        try:
            template_bytes = load_template_bytes()
        except FileNotFoundError:
            flash("No se encontro la plantilla Excel en el servidor.", "error")
            return redirect(url_for("proyectos"))

        return send_file(
            BytesIO(template_bytes),
            as_attachment=True,
            download_name="plantilla_riesgos.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.route("/proyectos/import", methods=["POST"])
    def proyectos_import():
        uploaded = request.files.get("file")
        if not uploaded or not uploaded.filename:
            flash("Selecciona un archivo Excel para importar.", "error")
            return redirect(url_for("proyectos"))

        try:
            workbook = parse_imported_workbook(uploaded.read())
        except Exception:
            flash("No se pudo leer el Excel.", "error")
            return redirect(url_for("proyectos"))

        for warning in workbook.warnings:
            flash(warning, "warning")

        project_payload = workbook.project or {
            "name": Path(uploaded.filename).stem or "Proyecto importado",
            "project_type": "",
            "company": "",
            "start_date": None,
            "end_date": None,
            "status": "Activo",
        }
        created_project = create_project(project_payload)

        for role in workbook.roles:
            create_project_role(created_project["id"], role)

        for deliverable in workbook.deliverables:
            create_project_deliverable(created_project["id"], deliverable)

        for schedule_item in workbook.schedule:
            create_project_schedule(created_project["id"], schedule_item)

        for threat in workbook.threats:
            created_threat = create_project_threat(created_project["id"], threat)
            if created_threat:
                threat.setdefault("id", created_threat["id"])

        threat_lookup = {threat.get("threat", ""): threat for threat in workbook.threats}
        for safeguard in workbook.safeguards:
            linked_threat = threat_lookup.get(safeguard.get("threat_name", ""))
            safeguard_payload = dict(safeguard)
            if linked_threat and linked_threat.get("id"):
                safeguard_payload["threat_id"] = linked_threat["id"]
            create_project_safeguard(created_project["id"], safeguard_payload)

        asset_map: dict[str, int] = {}
        for asset in workbook.assets:
            created_asset = create_asset(created_project["id"], asset)
            if created_asset:
                asset_map[_lookup_key(asset.get("name", ""))] = created_asset["id"]

        risk_map: dict[str, int] = {}
        for risk in workbook.risks:
            risk_payload = dict(risk)
            asset_name = risk_payload.pop("asset_name", "")
            asset_key = _lookup_key(asset_name)
            if asset_key and asset_key in asset_map:
                risk_payload["asset_id"] = asset_map[asset_key]
            created_risk = create_risk(created_project["id"], risk_payload)
            if created_risk:
                risk_map[_lookup_key(risk.get("name", ""))] = created_risk["id"]

        mitigation_count = 0
        for mitigation in workbook.mitigations:
            mitigation_payload = dict(mitigation)
            risk_name = mitigation_payload.pop("risk_name", "")
            mitigation_payload["risk_name"] = risk_name
            risk_key = _lookup_key(risk_name)
            if risk_key and risk_key in risk_map:
                mitigation_payload["risk_id"] = risk_map[risk_key]
            created_mitigation = create_mitigation(created_project["id"], mitigation_payload)
            if created_mitigation:
                mitigation_count += 1

        _set_selected_project(created_project["id"])
        flash(
            f"Importacion completada: 1 proyecto, {len(workbook.roles)} roles, {len(workbook.deliverables)} entregables, {len(workbook.schedule)} fases de cronograma, {len(workbook.threats)} amenazas, "
            f"{len(workbook.safeguards)} salvaguardas, {len(workbook.assets)} activos, "
            f"{len(workbook.risks)} riesgos y {mitigation_count} mitigaciones.",
            "success",
        )
        return redirect(url_for("proyectos"))

    @app.route("/activos", methods=["GET", "POST"])
    def activos():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            asset_id = request.form.get("asset_id", type=int)
            payload = {
                "name": request.form.get("name", ""),
                "type": request.form.get("type", ""),
                "value": request.form.get("value", 0),
            }
            if asset_id:
                if not get_asset(asset_id):
                    flash("El activo no existe.", "error")
                else:
                    update_asset(asset_id, payload)
                    flash("Activo actualizado.", "success")
            else:
                create_asset(project_id, payload)
                flash("Activo creado.", "success")
            return redirect(url_for("activos"))

        edit_id = request.args.get("edit", type=int)
        editing_asset = get_asset(edit_id) if edit_id else None
        return _render_assets(editing_asset=editing_asset)

    @app.route("/activos/delete/<int:asset_id>", methods=["POST"])
    def delete_asset_route(asset_id: int):
        if not get_asset(asset_id):
            flash("El activo no existe.", "error")
        else:
            delete_asset(asset_id)
            flash("Activo eliminado.", "success")
        return redirect(url_for("activos"))

    @app.route("/entregables", methods=["GET", "POST"])
    def entregables():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            deliverable_id = request.form.get("deliverable_id", type=int)
            payload = {
                "deliverable": request.form.get("deliverable", ""),
                "description": request.form.get("description", ""),
                "main_responsible": request.form.get("main_responsible", ""),
                "status": request.form.get("status", "Pendiente"),
            }
            if deliverable_id:
                if not get_project_deliverable(deliverable_id):
                    flash("El entregable no existe.", "error")
                else:
                    update_project_deliverable(deliverable_id, payload)
                    flash("Entregable actualizado.", "success")
            else:
                create_project_deliverable(project_id, payload)
                flash("Entregable creado.", "success")
            return redirect(url_for("entregables"))

        edit_id = request.args.get("edit", type=int)
        editing_deliverable = get_project_deliverable(edit_id) if edit_id else None
        return _render_deliverables(editing_deliverable=editing_deliverable)

    @app.route("/entregables/delete/<int:deliverable_id>", methods=["POST"])
    def delete_deliverable_route(deliverable_id: int):
        if not get_project_deliverable(deliverable_id):
            flash("El entregable no existe.", "error")
        else:
            delete_project_deliverable(deliverable_id)
            flash("Entregable eliminado.", "success")
        return redirect(url_for("entregables"))

    @app.route("/cronograma", methods=["GET", "POST"])
    def cronograma():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        if request.method == "POST":
            schedule_id = request.form.get("schedule_id", type=int)
            payload = {
                "phase": request.form.get("phase", ""),
                "risks": request.form.get("risks", ""),
                "control": request.form.get("control", ""),
                "start_period": request.form.get("start_period", type=int) or 1,
                "end_period": request.form.get("end_period", type=int) or 1,
            }
            if payload["end_period"] < payload["start_period"]:
                payload["end_period"] = payload["start_period"]
            if schedule_id:
                if not get_project_schedule(schedule_id):
                    flash("La fase del cronograma no existe.", "error")
                else:
                    update_project_schedule(schedule_id, payload)
                    flash("Fase del cronograma actualizada.", "success")
            else:
                create_project_schedule(project_id, payload)
                flash("Fase del cronograma creada.", "success")
            return redirect(url_for("cronograma"))
        return _render_cronograma()

    @app.route("/cronograma/delete/<int:schedule_id>", methods=["POST"])
    def delete_schedule_route(schedule_id: int):
        if not get_project_schedule(schedule_id):
            flash("La fase del cronograma no existe.", "error")
        else:
            delete_project_schedule(schedule_id)
            flash("Fase del cronograma eliminada.", "success")
        return redirect(url_for("cronograma"))

    @app.route("/proyecto/pdf")
    def pdf_proyecto():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        return _render_pdf_project()

    @app.route("/proyecto/pdf/download")
    def pdf_proyecto_download():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        context = _project_report_context(project_id)
        pdf_bytes = _build_project_pdf_clean(context)
        project = context["project"] or {}
        filename = f"proyecto_{project.get('code') or project_id}.pdf"
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/pdf",
        )

    @app.route("/riesgos", methods=["GET", "POST"])
    def riesgos():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            risk_id = request.form.get("risk_id", type=int)
            payload = {
                "asset_id": request.form.get("asset_id", type=int),
                "asset_name": request.form.get("asset_name", ""),
                "name": request.form.get("name", ""),
                "description": request.form.get("description", ""),
                "cause": request.form.get("cause", ""),
                "consequence": request.form.get("consequence", ""),
                "probability": request.form.get("probability", 0),
                "impact": request.form.get("impact", 0),
                "level": request.form.get("level", ""),
                "horizon": request.form.get("horizon", ""),
                "owner": request.form.get("owner", ""),
                "status": request.form.get("status", "Identificado"),
                "strategy": request.form.get("strategy", "Mitigar"),
            }
            if risk_id:
                if not get_risk(risk_id):
                    flash("El riesgo no existe.", "error")
                else:
                    update_risk(risk_id, payload)
                    flash("Riesgo actualizado.", "success")
            else:
                create_risk(project_id, payload)
                flash("Riesgo creado.", "success")
            return redirect(url_for("riesgos"))

        edit_id = request.args.get("edit", type=int)
        editing_risk = get_risk(edit_id) if edit_id else None
        return _render_risks(editing_risk=editing_risk)

    @app.route("/riesgos/delete/<int:risk_id>", methods=["POST"])
    def delete_risk_route(risk_id: int):
        if not get_risk(risk_id):
            flash("El riesgo no existe.", "error")
        else:
            delete_risk(risk_id)
            flash("Riesgo eliminado.", "success")
        return redirect(url_for("riesgos"))

    @app.route("/riesgos/roles", methods=["GET", "POST"])
    def roles():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            role_id = request.form.get("role_id", type=int)
            payload = {
                "role": request.form.get("role", ""),
                "acquisition_type": request.form.get("acquisition_type", ""),
                "risk_participation": request.form.get("risk_participation", ""),
            }
            if role_id:
                if not get_project_role(role_id):
                    flash("El rol no existe.", "error")
                else:
                    update_project_role(role_id, payload)
                    flash("Rol actualizado.", "success")
            else:
                create_project_role(project_id, payload)
                flash("Rol creado.", "success")
            return redirect(url_for("roles"))

        edit_id = request.args.get("edit", type=int)
        editing_role = get_project_role(edit_id) if edit_id else None
        return _render_roles(editing_role=editing_role)

    @app.route("/riesgos/roles/delete/<int:role_id>", methods=["POST"])
    def delete_role_route(role_id: int):
        if not get_project_role(role_id):
            flash("El rol no existe.", "error")
        else:
            delete_project_role(role_id)
            flash("Rol eliminado.", "success")
        return redirect(url_for("roles"))

    @app.route("/riesgos/amenazas", methods=["GET", "POST"])
    def amenazas():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            threat_id = request.form.get("threat_id", type=int)
            payload = {
                "threat": request.form.get("threat", ""),
                "affected_asset": request.form.get("affected_asset", ""),
                "example": request.form.get("example", ""),
            }
            if threat_id:
                if not get_project_threat(threat_id):
                    flash("La amenaza no existe.", "error")
                else:
                    update_project_threat(threat_id, payload)
                    flash("Amenaza actualizada.", "success")
            else:
                create_project_threat(project_id, payload)
                flash("Amenaza creada.", "success")
            return redirect(url_for("amenazas"))

        edit_id = request.args.get("edit", type=int)
        editing_threat = get_project_threat(edit_id) if edit_id else None
        return _render_threats(editing_threat=editing_threat)

    @app.route("/riesgos/amenazas/delete/<int:threat_id>", methods=["POST"])
    def delete_threat_route(threat_id: int):
        if not get_project_threat(threat_id):
            flash("La amenaza no existe.", "error")
        else:
            delete_project_threat(threat_id)
            flash("Amenaza eliminada.", "success")
        return redirect(url_for("amenazas"))

    @app.route("/riesgos/salvaguardas", methods=["POST"])
    def salvaguardas():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        safeguard_id = request.form.get("safeguard_id", type=int)
        payload = {
            "threat_id": request.form.get("threat_id", type=int),
            "threat_name": request.form.get("threat_name", ""),
            "safeguard": request.form.get("safeguard", ""),
        }
        if safeguard_id:
            if not get_project_safeguard(safeguard_id):
                flash("La salvaguarda no existe.", "error")
            else:
                update_project_safeguard(safeguard_id, payload)
                flash("Salvaguarda actualizada.", "success")
        else:
            create_project_safeguard(project_id, payload)
            flash("Salvaguarda creada.", "success")
        return redirect(url_for("amenazas"))

    @app.route("/riesgos/salvaguardas/delete/<int:safeguard_id>", methods=["POST"])
    def delete_safeguard_route(safeguard_id: int):
        if not get_project_safeguard(safeguard_id):
            flash("La salvaguarda no existe.", "error")
        else:
            delete_project_safeguard(safeguard_id)
            flash("Salvaguarda eliminada.", "success")
        return redirect(url_for("amenazas"))

    @app.route("/matriz")
    def matriz():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        return _render_matrix()

    @app.route("/mitigacion", methods=["GET", "POST"])
    def mitigacion():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))

        if request.method == "POST":
            mitigation_id = request.form.get("mitigation_id", type=int)
            payload = {
                "risk_id": request.form.get("risk_id", type=int),
                "risk_name": request.form.get("risk_name", ""),
                "preventive_action": request.form.get("preventive_action", ""),
                "corrective_action": request.form.get("corrective_action", ""),
                "trigger": request.form.get("trigger", ""),
            }
            if mitigation_id:
                if not get_mitigation(mitigation_id):
                    flash("La mitigacion no existe.", "error")
                else:
                    update_mitigation(mitigation_id, payload)
                    flash("Mitigacion actualizada.", "success")
            else:
                create_mitigation(project_id, payload)
                flash("Mitigacion creada.", "success")
            return redirect(url_for("mitigacion"))

        edit_id = request.args.get("edit", type=int)
        editing_mitigation = get_mitigation(edit_id) if edit_id else None
        return _render_mitigations(editing_mitigation=editing_mitigation)

    @app.route("/mitigacion/delete/<int:mitigation_id>", methods=["POST"])
    def delete_mitigation_route(mitigation_id: int):
        if not get_mitigation(mitigation_id):
            flash("La mitigacion no existe.", "error")
        else:
            delete_mitigation(mitigation_id)
            flash("Mitigacion eliminada.", "success")
        return redirect(url_for("mitigacion"))

    @app.route("/reportes")
    def reportes():
        project_id = _require_project()
        if not project_id:
            return redirect(url_for("proyectos"))
        return _render_reports()

    @app.cli.command("reset-db")
    def reset_db_command() -> None:
        reset_db()
        print("SQLite database reset.")

    return app


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
