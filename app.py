import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import threading
import time

from apscheduler.schedulers.background import BackgroundScheduler

from app_config import load_config
from email_alerter import EmailAlertManager
from scheduler import CycleOrchestrator
from api_routes import create_app


def setup_logging():
    """Configura logging su console e file rotativo."""
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(log_dir, "plan_monitor.log")

    logger = logging.getLogger("PlanMonitor")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    logger.addHandler(console)

    file_handler = RotatingFileHandler(
        log_file, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


def main():
    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("Production Plan Monitor - Avvio")
    logger.info("=" * 60)

    config = load_config()
    logger.info("Configurazione caricata: polling=%d min, orario=%s-%s, porta=%d",
                config.polling.interval_minutes,
                config.workday.start.strftime("%H:%M"),
                config.workday.end.strftime("%H:%M"),
                config.server.port)

    email_alerter = EmailAlertManager(config)
    orchestrator = CycleOrchestrator(config, email_alerter)

    app = create_app(config, orchestrator)

    # Setup APScheduler
    scheduler = BackgroundScheduler()
    scheduler.add_job(
        orchestrator.run_cycle,
        'interval',
        minutes=config.polling.interval_minutes,
        id='production_monitor_cycle',
        name='Production Monitor Cycle',
        max_instances=1,
        coalesce=True
    )
    scheduler.start()
    logger.info("Scheduler avviato: ciclo ogni %d minuti", config.polling.interval_minutes)

    # Primo ciclo in background
    def _initial_cycle():
        time.sleep(3)
        logger.info("Esecuzione primo ciclo iniziale...")
        orchestrator.run_cycle(force=True)

    t = threading.Thread(target=_initial_cycle, daemon=True)
    t.start()

    # Avvia Flask
    try:
        logger.info("Server Flask avviato su http://%s:%d", config.server.host, config.server.port)
        app.run(
            host=config.server.host,
            port=config.server.port,
            debug=False,
            use_reloader=False
        )
    finally:
        scheduler.shutdown(wait=False)
        logger.info("Production Plan Monitor - Arresto")


if __name__ == "__main__":
    main()
