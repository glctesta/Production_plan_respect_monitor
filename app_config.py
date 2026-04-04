import yaml
import os
from datetime import time, date
from dataclasses import dataclass, field
from typing import List


@dataclass
class PlanningConfig:
    folder: str = "T:\\Planning"
    sheet: str = "Last Phase"


@dataclass
class WorkdayConfig:
    start: time = field(default_factory=lambda: time(7, 30))
    end: time = field(default_factory=lambda: time(23, 30))
    total_minutes: int = 900


@dataclass
class PollingConfig:
    interval_minutes: int = 10


@dataclass
class ThresholdsConfig:
    red_deficit: int = 10


@dataclass
class EmailConfig:
    enabled: bool = True
    settings_attribute: str = "sys_email_planning_warning"
    yellow_cooldown_minutes: int = 120
    red_cooldown_minutes: int = 60


@dataclass
class HolidaysConfig:
    dates: List[date] = field(default_factory=list)

    def is_holiday(self, d: date) -> bool:
        return d in self.dates

    def is_working_day(self, d: date) -> bool:
        """Ritorna True se il giorno e' lavorativo (lun-ven e non festivo)."""
        if d.weekday() >= 5:  # 5=sabato, 6=domenica
            return False
        return not self.is_holiday(d)


@dataclass
class UIConfig:
    enable_blinking_alerts: bool = True


@dataclass
class ServerConfig:
    host: str = "0.0.0.0"
    port: int = 8085


@dataclass
class AppConfig:
    planning: PlanningConfig = field(default_factory=PlanningConfig)
    workday: WorkdayConfig = field(default_factory=WorkdayConfig)
    polling: PollingConfig = field(default_factory=PollingConfig)
    thresholds: ThresholdsConfig = field(default_factory=ThresholdsConfig)
    email: EmailConfig = field(default_factory=EmailConfig)
    holidays: HolidaysConfig = field(default_factory=HolidaysConfig)
    ui: UIConfig = field(default_factory=UIConfig)
    server: ServerConfig = field(default_factory=ServerConfig)


def _parse_time(value: str) -> time:
    parts = value.strip().split(":")
    return time(int(parts[0]), int(parts[1]))


def load_config(path: str = None) -> AppConfig:
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")

    config = AppConfig()

    if not os.path.exists(path):
        return config

    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    # Planning
    p = raw.get("planning", {})
    config.planning.folder = p.get("folder", config.planning.folder)
    config.planning.sheet = p.get("sheet", config.planning.sheet)

    # Workday
    w = raw.get("workday", {})
    if "start" in w:
        config.workday.start = _parse_time(str(w["start"]))
    if "end" in w:
        config.workday.end = _parse_time(str(w["end"]))
    start_min = config.workday.start.hour * 60 + config.workday.start.minute
    end_min = config.workday.end.hour * 60 + config.workday.end.minute
    config.workday.total_minutes = end_min - start_min

    # Polling
    pl = raw.get("polling", {})
    config.polling.interval_minutes = pl.get("interval_minutes", config.polling.interval_minutes)

    # Thresholds
    t = raw.get("thresholds", {})
    config.thresholds.red_deficit = t.get("red_deficit", config.thresholds.red_deficit)

    # Email
    e = raw.get("email", {})
    config.email.enabled = e.get("enabled", config.email.enabled)
    config.email.settings_attribute = e.get("settings_attribute", config.email.settings_attribute)
    config.email.yellow_cooldown_minutes = e.get("yellow_cooldown_minutes", config.email.yellow_cooldown_minutes)
    config.email.red_cooldown_minutes = e.get("red_cooldown_minutes", config.email.red_cooldown_minutes)

    # Holidays
    h = raw.get("holidays", {})
    holiday_dates = h.get("dates", [])
    parsed_holidays = []
    for hd in holiday_dates:
        try:
            if isinstance(hd, date):
                parsed_holidays.append(hd)
            else:
                parsed_holidays.append(date.fromisoformat(str(hd).strip()))
        except (ValueError, TypeError):
            pass  # Ignora date malformate
    config.holidays.dates = parsed_holidays

    # UI
    u = raw.get("ui", {})
    config.ui.enable_blinking_alerts = u.get("enable_blinking_alerts", config.ui.enable_blinking_alerts)

    # Server
    s = raw.get("server", {})
    config.server.host = s.get("host", config.server.host)
    config.server.port = s.get("port", config.server.port)

    return config
