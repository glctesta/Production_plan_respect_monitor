import logging
from datetime import datetime, date
from typing import List, Tuple, Optional

from app_config import AppConfig
from monitor_engine import MonitorRow
from utils import get_email_recipients, send_email

logger = logging.getLogger("PlanMonitor")


class EmailAlertManager:
    def __init__(self, config: AppConfig):
        self.config = config
        self._last_yellow_email: Optional[datetime] = None
        self._last_red_email: Optional[datetime] = None
        self._previous_red_count: int = 0
        self._today: Optional[date] = None

    def _reset_if_new_day(self):
        today = date.today()
        if self._today != today:
            self._today = today
            self._last_yellow_email = None
            self._last_red_email = None
            self._previous_red_count = 0
            logger.info("Email alert state reset per nuovo giorno: %s", today)

    def should_send_email(self, rows: List[MonitorRow]) -> Tuple[bool, int]:
        """
        Determina se inviare email e il livello di severita.
        Livelli: 0=nessuna, 1=solo gialli, 2=rossi presenti, 3=rossi in aumento.
        Rispetta cooldown: giallo max 1 ogni 2h, rosso max 1 ogni 1h.
        """
        self._reset_if_new_day()

        if not self.config.email.enabled:
            return (False, 0)

        yellow_count = sum(1 for r in rows if r.status_color == "yellow" and not r.is_out_of_plan)
        red_count = sum(1 for r in rows if r.status_color == "red" and not r.is_out_of_plan)
        out_of_plan_count = sum(1 for r in rows if r.is_out_of_plan)

        # Include out-of-plan come "rossi" ai fini dell'email
        total_red = red_count + out_of_plan_count

        if yellow_count == 0 and total_red == 0:
            return (False, 0)

        now = datetime.now()

        # Determina severita
        if total_red > 0:
            if total_red > self._previous_red_count and self._previous_red_count > 0:
                severity = 3  # Rossi in aumento
            else:
                severity = 2  # Rossi presenti
        else:
            severity = 1  # Solo gialli

        # Controlla cooldown
        if severity >= 2:
            # Rosso: cooldown 1 ora
            if self._last_red_email:
                elapsed = (now - self._last_red_email).total_seconds() / 60.0
                if elapsed < self.config.email.red_cooldown_minutes:
                    logger.info("Email rosso saltata: cooldown attivo (%.0f min trascorsi, richiesti %d)",
                                elapsed, self.config.email.red_cooldown_minutes)
                    return (False, 0)
        else:
            # Giallo: cooldown 2 ore
            if self._last_yellow_email:
                elapsed = (now - self._last_yellow_email).total_seconds() / 60.0
                if elapsed < self.config.email.yellow_cooldown_minutes:
                    logger.info("Email giallo saltata: cooldown attivo (%.0f min trascorsi, richiesti %d)",
                                elapsed, self.config.email.yellow_cooldown_minutes)
                    return (False, 0)

        return (True, severity)

    def _build_email_html(self, rows: List[MonitorRow], severity: int,
                          summary: dict, excel_file: str) -> Tuple[str, str]:
        """Genera subject e body HTML dell'email."""
        # Subject per livello
        if severity == 3:
            subject = "URGENT: Production Delays Are WORSENING - Immediate Action Required"
        elif severity == 2:
            subject = "WARNING: Critical Production Delays Detected"
        else:
            subject = "ATTENTION: Production Lines Behind Schedule"

        # Intro per livello
        if severity == 3:
            intro = (
                "<p style='color: #c0392b; font-size: 16px; font-weight: bold;'>"
                "CRITICAL ESCALATION: The number of production lines with critical delays "
                "has INCREASED since the last check. This situation demands your immediate attention "
                "and corrective action. Failure to intervene NOW will result in missed daily targets.</p>"
            )
        elif severity == 2:
            intro = (
                "<p style='color: #e67e22; font-size: 15px; font-weight: bold;'>"
                "Several production lines are significantly behind schedule and are projected "
                "to MISS their daily targets. Review the details below and take corrective action immediately.</p>"
            )
        else:
            intro = (
                "<p style='color: #f39c12;'>"
                "Some production lines are falling slightly behind the planned schedule. "
                "Please monitor closely and adjust resources if needed.</p>"
            )

        # Tabella righe problematiche
        problem_rows = [r for r in rows if r.status_color in ("yellow", "red") or r.is_out_of_plan]

        table_rows = ""
        for r in problem_rows:
            if r.is_out_of_plan:
                bg = "#e74c3c"
                fg = "white"
                status = "OUT OF PLAN"
            elif r.status_color == "red":
                bg = "#fadbd8"
                fg = "#333"
                status = "CRITICAL"
            else:
                bg = "#fdebd0"
                fg = "#333"
                status = "WARNING"

            # Star indicator for context
            star_html = ""
            if r.context_star == "yellow":
                star_html = " &#9733;"  # yellow star
                status += " (scheduled upcoming)"
            elif r.context_star == "blue":
                star_html = " &#9733;"  # blue star
                status += " (DELAYED)"

            # Context note as extra info
            context_row = ""
            if r.context_note:
                context_row = (
                    f"<tr style='background-color: {bg}; color: {fg}; font-size: 11px;'>"
                    f"<td colspan='7' style='padding: 2px 8px 6px 20px; border: 1px solid #ddd; "
                    f"font-style: italic;'>&#8627; {r.context_note}</td>"
                    f"</tr>"
                )

            table_rows += (
                f"<tr style='background-color: {bg}; color: {fg};'>"
                f"<td style='padding: 8px; border: 1px solid #ddd;'>{r.order_number}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd;'>{r.product_code}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd;'>{r.phase}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>{r.planned_qty_day}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>{r.qty_done}{star_html}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>{r.projected_deficit}</td>"
                f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;'>{status}</td>"
                f"</tr>"
                f"{context_row}"
            )

        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <h2 style="color: #2c3e50;">Production Plan Monitoring Alert</h2>
            {intro}
            <p><strong>Summary:</strong>
                Green: {summary.get('green', 0)} |
                <span style="color: #f39c12;">Yellow: {summary.get('yellow', 0)}</span> |
                <span style="color: #e74c3c;">Red: {summary.get('red', 0)}</span> |
                <span style="color: #c0392b; font-weight: bold;">Out of Plan: {summary.get('out_of_plan', 0)}</span>
            </p>
            <p><strong>Source file:</strong> {excel_file}</p>
            <p><strong>Timestamp:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

            <table style="border-collapse: collapse; width: 100%; margin: 15px 0;">
                <thead>
                    <tr style="background-color: #2c3e50; color: white;">
                        <th style="padding: 10px; border: 1px solid #ddd;">Order</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Product</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Phase</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Planned</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Done</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Projected Deficit</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>

            <hr style="margin: 20px 0;"/>
            <p style="color: #666; font-size: 12px;">
                This is an automated alert from the Production Plan Monitoring System.<br/>
                Do not reply to this email.
            </p>
        </body>
        </html>
        """
        return subject, body

    def send_alerts(self, conn, rows: List[MonitorRow], summary: dict, excel_file: str) -> bool:
        """Invia email di alert se necessario, rispettando i cooldown."""
        should_send, severity = self.should_send_email(rows)
        if not should_send:
            return False

        try:
            recipients = get_email_recipients(conn, self.config.email.settings_attribute)
            if not recipients:
                logger.warning("Nessun destinatario trovato per attributo '%s'",
                               self.config.email.settings_attribute)
                return False

            subject, body = self._build_email_html(rows, severity, summary, excel_file)

            send_email(
                recipients=recipients,
                subject=subject,
                body=body,
                is_html=True
            )

            # Aggiorna stato cooldown
            now = datetime.now()
            if severity >= 2:
                self._last_red_email = now
            else:
                self._last_yellow_email = now

            red_count = sum(1 for r in rows if (r.status_color == "red" and not r.is_out_of_plan))
            out_of_plan = sum(1 for r in rows if r.is_out_of_plan)
            self._previous_red_count = red_count + out_of_plan

            logger.info("Email alert inviata: severity=%d, destinatari=%d", severity, len(recipients))
            return True

        except Exception as e:
            logger.error("Errore invio email alert: %s", e)
            return False
