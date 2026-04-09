import os
import threading
from flask import Flask, render_template, jsonify, request

from app_config import AppConfig
from scheduler import CycleOrchestrator


def create_app(config: AppConfig, orchestrator: CycleOrchestrator) -> Flask:
    base_dir = os.path.dirname(os.path.abspath(__file__))

    app = Flask(
        __name__,
        static_folder=os.path.join(base_dir, "static"),
        template_folder=os.path.join(base_dir, "templates")
    )

    @app.route("/")
    def dashboard():
        return render_template("index.html", config=config)

    @app.route("/api/status")
    def api_status():
        return jsonify(orchestrator.get_status())

    @app.route("/api/run-now", methods=["POST"])
    def api_run_now():
        def _run():
            orchestrator.run_cycle(force=True)
        t = threading.Thread(target=_run, daemon=True)
        t.start()
        return jsonify({"status": "started", "message": "Ciclo manuale avviato"})

    @app.route("/api/config")
    def api_config():
        return jsonify({
            "planning_folder": config.planning.folder,
            "planning_sheet": config.planning.sheet,
            "workday_start": config.workday.start.strftime("%H:%M"),
            "workday_end": config.workday.end.strftime("%H:%M"),
            "poll_minutes": config.polling.interval_minutes,
            "red_deficit_threshold": config.thresholds.red_deficit,
            "email_enabled": config.email.enabled,
            "yellow_cooldown_minutes": config.email.yellow_cooldown_minutes,
            "red_cooldown_minutes": config.email.red_cooldown_minutes,
            "blinking_enabled": config.ui.enable_blinking_alerts
        })

    @app.route("/api/alert-response", methods=["POST"])
    def api_alert_response():
        """Endpoint per inserire risposte/giustificazioni operatori per un alert."""
        data = request.get_json()
        if not data:
            return jsonify({"error": "JSON body required"}), 400

        alert_id = data.get("alert_id")
        operator_name = data.get("operator_name", "").strip()
        response_text = data.get("response", "").strip()

        if not alert_id or not operator_name or not response_text:
            return jsonify({"error": "alert_id, operator_name and response are required"}), 400

        try:
            from db_queries import get_db_connection, insert_plan_alert_response
            db = get_db_connection()
            conn = db.connect()
            success = insert_plan_alert_response(conn, int(alert_id), operator_name, response_text)
            db.disconnect()
            if success:
                return jsonify({"status": "ok", "message": "Response saved"})
            else:
                return jsonify({"error": "Failed to save response"}), 500
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/api/output-summary")
    def api_output_summary():
        return jsonify(orchestrator.get_output_status())

    @app.route("/api/output-config")
    def api_output_config():
        return jsonify(orchestrator.output_config)

    @app.route("/api/health")
    def api_health():
        folder_ok = os.path.isdir(config.planning.folder)
        db_ok = False
        try:
            from db_queries import get_db_connection
            db = get_db_connection()
            conn = db.connect()
            db_ok = True
            db.disconnect()
        except Exception:
            pass
        return jsonify({
            "status": "ok" if (folder_ok and db_ok) else "degraded",
            "db_connected": db_ok,
            "planning_folder_accessible": folder_ok
        })

    return app
