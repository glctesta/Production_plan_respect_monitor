import os
import threading
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app_config import AppConfig
from scheduler import CycleOrchestrator


def create_app(config: AppConfig, orchestrator: CycleOrchestrator) -> FastAPI:
    app = FastAPI(title="Production Plan Monitor", version="1.0.0")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    static_dir = os.path.join(base_dir, "static")
    templates_dir = os.path.join(base_dir, "templates")

    os.makedirs(static_dir, exist_ok=True)
    os.makedirs(templates_dir, exist_ok=True)

    app.mount("/static", StaticFiles(directory=static_dir), name="static")
    templates = Jinja2Templates(directory=templates_dir)

    @app.get("/", response_class=HTMLResponse)
    async def dashboard(request: Request):
        return templates.TemplateResponse(
            request=request,
            name="index.html",
            context={"config": config}
        )

    @app.get("/api/status")
    async def api_status():
        return JSONResponse(content=orchestrator.get_status())

    @app.post("/api/run-now")
    async def api_run_now():
        def _run():
            orchestrator.run_cycle(force=True)
        t = threading.Thread(target=_run, daemon=True)
        t.start()
        return JSONResponse(content={"status": "started", "message": "Ciclo manuale avviato"})

    @app.get("/api/config")
    async def api_config():
        return JSONResponse(content={
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

    @app.get("/api/health")
    async def api_health():
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
        return JSONResponse(content={
            "status": "ok" if (folder_ok and db_ok) else "degraded",
            "db_connected": db_ok,
            "planning_folder_accessible": folder_ok
        })

    return app
