from __future__ import annotations

import os
from io import BytesIO
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

from db import (
    close_db,
    create_asset,
    create_mitigation,
    create_project,
    create_risk,
    delete_asset,
    delete_mitigation,
    delete_project,
    delete_risk,
    get_asset,
    get_mitigation,
    get_project,
    get_risk,
    init_db,
    list_assets,
    list_assets_for_select,
    list_mitigations,
    list_projects,
    list_risk_names,
    list_risks,
    report_summary,
    reset_db,
    update_asset,
    update_mitigation,
    update_project,
    update_risk,
)
from excel_utils import load_template_bytes, parse_imported_workbook


BASE_DIR = Path(__file__).resolve().parent


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
            "activos",
            "riesgos",
            "matriz",
            "mitigacion",
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
        if not project_id:
            flash("Crea o selecciona un proyecto primero.", "warning")
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

    def _render_matrix():
        project_id = _selected_project_id()
        risks = list_risks(project_id)
        return render_template(
            "matriz.html",
            title="Matriz",
            active_page="matriz",
            matrix_risks=risks,
        )

    def _render_mitigations(editing_mitigation: dict[str, Any] | None = None):
        project_id = _selected_project_id()
        return render_template(
            "mitigacion.html",
            title="Mitigacion",
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

        asset_map: dict[str, int] = {}
        for asset in workbook.assets:
            created_asset = create_asset(created_project["id"], asset)
            if created_asset:
                asset_map[asset.get("name", "")] = created_asset["id"]

        risk_map: dict[str, int] = {}
        for risk in workbook.risks:
            risk_payload = dict(risk)
            asset_name = risk_payload.pop("asset_name", "")
            if asset_name and asset_name in asset_map:
                risk_payload["asset_id"] = asset_map[asset_name]
            created_risk = create_risk(created_project["id"], risk_payload)
            if created_risk:
                risk_map[risk.get("name", "")] = created_risk["id"]

        mitigation_count = 0
        for mitigation in workbook.mitigations:
            mitigation_payload = dict(mitigation)
            risk_name = mitigation_payload.pop("risk_name", "")
            if risk_name and risk_name in risk_map:
                mitigation_payload["risk_id"] = risk_map[risk_name]
            created_mitigation = create_mitigation(created_project["id"], mitigation_payload)
            if created_mitigation:
                mitigation_count += 1

        _set_selected_project(created_project["id"])
        flash(
            f"Importacion completada: 1 proyecto, {len(workbook.assets)} activos, "
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
                "owner": request.form.get("owner", ""),
                "value": request.form.get("value", 0),
                "status": request.form.get("status", "Activo"),
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
                "owner": request.form.get("owner", ""),
                "start_date": request.form.get("start_date"),
                "end_date": request.form.get("end_date"),
                "resources": request.form.get("resources", ""),
                "status": request.form.get("status", "Planificado"),
                "evidence": request.form.get("evidence", ""),
                "strategy": request.form.get("strategy", "Mitigar"),
                "progress": request.form.get("progress", 0),
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
