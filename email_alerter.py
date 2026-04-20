import logging
import os
from datetime import datetime, date, time as dtime
from typing import List, Tuple, Optional

from app_config import AppConfig
from monitor_engine import MonitorRow
from utils import (get_email_recipients, send_email,
                   is_visible_phase, display_phase_name)

logger = logging.getLogger("PlanMonitor")

# Percorso logo per email inline
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "Logo.png")


# ── Schedule email warning: 1 sola mail per turno, verso la meta' ────────────
# Turno 1: 07:30 -> 15:30 (meta' ~ 11:30); finestra utile 11:00-12:30.
# Turno 2: 15:31 -> 23:30 (meta' ~ 19:30); finestra utile 19:00-20:30.
SHIFT1_WINDOW = (dtime(11, 0), dtime(12, 30))
SHIFT2_WINDOW = (dtime(19, 0), dtime(20, 30))


def _current_shift(now: datetime) -> Optional[int]:
    """Ritorna 1 se siamo nella finestra utile del turno mattina, 2 per la sera,
    None se fuori."""
    t = now.time()
    if SHIFT1_WINDOW[0] <= t <= SHIFT1_WINDOW[1]:
        return 1
    if SHIFT2_WINDOW[0] <= t <= SHIFT2_WINDOW[1]:
        return 2
    return None


class EmailAlertManager:
    def __init__(self, config: AppConfig):
        self.config = config
        self._today: Optional[date] = None
        self._sent_shifts: set = set()          # shift numeri {1,2} gia' inviati oggi
        self._previous_red_count: int = 0
        self._adjustment_email_sent_today: bool = False

    def _reset_if_new_day(self):
        today = date.today()
        if self._today != today:
            self._today = today
            self._sent_shifts = set()
            self._previous_red_count = 0
            self._adjustment_email_sent_today = False
            logger.info("Email alert state reset per nuovo giorno: %s", today)

    def should_send_email(self, rows: List[MonitorRow]) -> Tuple[bool, int]:
        """
        Politica nuova: al massimo 1 email di warning per turno.
          - finestra turno 1 (11:00-12:30)
          - finestra turno 2 (19:00-20:30)
        Se ci sono alert sulle sole fasi visibili (AOI/PTHM) e la relativa email
        del turno non e' ancora stata mandata oggi, si invia.
        Livelli: 0=nessuna, 1=solo gialli, 2=rossi presenti, 3=rossi in aumento.
        """
        self._reset_if_new_day()

        if not self.config.email.enabled:
            return (False, 0)

        now = datetime.now()
        shift = _current_shift(now)
        if shift is None:
            return (False, 0)

        if shift in self._sent_shifts:
            logger.info("Email turno %d: gia' inviata oggi, skip.", shift)
            return (False, 0)

        # Filtro solo sulle fasi visibili
        filtered = [r for r in rows if is_visible_phase(r.phase)]

        yellow_count = sum(1 for r in filtered if r.status_color == "yellow" and not r.is_out_of_plan)
        red_count = sum(1 for r in filtered if r.status_color == "red" and not r.is_out_of_plan)
        out_of_plan_count = sum(1 for r in filtered if r.is_out_of_plan)
        total_red = red_count + out_of_plan_count

        if yellow_count == 0 and total_red == 0:
            return (False, 0)

        if total_red > 0:
            if total_red > self._previous_red_count and self._previous_red_count > 0:
                severity = 3
            else:
                severity = 2
        else:
            severity = 1

        return (True, severity)

    @staticmethod
    def _summary_from_visible(rows: List[MonitorRow]) -> dict:
        """Riepilogo verdi/gialli/rossi/fuori piano calcolato sulle sole fasi visibili."""
        vis = [r for r in rows if is_visible_phase(r.phase)]
        green = sum(1 for r in vis if r.status_color == "green" and not r.is_out_of_plan)
        yellow = sum(1 for r in vis if r.status_color == "yellow" and not r.is_out_of_plan)
        red = sum(1 for r in vis if r.status_color == "red" and not r.is_out_of_plan)
        out_of_plan = sum(1 for r in vis if r.is_out_of_plan)
        return {"green": green, "yellow": yellow, "red": red, "out_of_plan": out_of_plan}

    def _build_email_html(self, rows: List[MonitorRow], severity: int,
                          summary: dict, excel_file: str) -> Tuple[str, str]:
        """Genera subject e body HTML dell'email."""
        # Summary ri-calcolato sulle sole fasi visibili
        visible_summary = self._summary_from_visible(rows)
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

        # Tabella righe problematiche: solo fasi visibili (AOI -> SMT, PTHM)
        problem_rows = [
            r for r in rows
            if is_visible_phase(r.phase)
            and (r.status_color in ("yellow", "red") or r.is_out_of_plan)
        ]

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
                f"<td style='padding: 8px; border: 1px solid #ddd;'>{display_phase_name(r.phase)}</td>"
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
            <p><strong>Summary</strong> (fasi SMT / PTHM):
                Green: {visible_summary.get('green', 0)} |
                <span style="color: #f39c12;">Yellow: {visible_summary.get('yellow', 0)}</span> |
                <span style="color: #e74c3c;">Red: {visible_summary.get('red', 0)}</span> |
                <span style="color: #c0392b; font-weight: bold;">Out of Plan: {visible_summary.get('out_of_plan', 0)}</span>
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
        """
        Invia email di warning, al massimo UNA volta per turno nelle finestre
        intorno alla meta' del turno (mattina ~11:30, pomeriggio ~19:30).
        Contenuto limitato alle fasi visibili (AOI -> SMT, PTHM).
        """
        should_send, severity = self.should_send_email(rows)
        if not should_send:
            return False

        now = datetime.now()
        shift = _current_shift(now)
        if shift is None:
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

            # Marca il turno come "inviato" per non re-inviare nella stessa finestra
            self._sent_shifts.add(shift)

            visible = [r for r in rows if is_visible_phase(r.phase)]
            red_count = sum(1 for r in visible if (r.status_color == "red" and not r.is_out_of_plan))
            out_of_plan = sum(1 for r in visible if r.is_out_of_plan)
            self._previous_red_count = red_count + out_of_plan

            logger.info("Email alert turno %d inviata: severity=%d, destinatari=%d",
                        shift, severity, len(recipients))
            return True

        except Exception as e:
            logger.error("Errore invio email alert: %s", e)
            return False

    def send_qty_adjustment_email(self, conn, adjusted_rows: List[MonitorRow],
                                   excel_file: str) -> bool:
        """
        Invia email di notifica aggiustamento quantita' (1 volta al giorno).
        adjusted_rows: lista di MonitorRow con qty_adjusted=True. Contenuto
        limitato alle fasi visibili (AOI -> SMT, PTHM).
        """
        self._reset_if_new_day()

        if self._adjustment_email_sent_today:
            logger.info("Email aggiustamento qty gia' inviata oggi, skip")
            return False

        # Filtra per fasi visibili
        adjusted_rows = [r for r in adjusted_rows if is_visible_phase(r.phase)]
        if not adjusted_rows:
            return False

        try:
            recipients = get_email_recipients(conn, self.config.email.settings_attribute)
            if not recipients:
                logger.warning("Nessun destinatario per email aggiustamento qty")
                return False

            subject = "NOTICE: Planning Quantities Adjusted Based on Traceability Data"

            # Costruisci righe tabella
            table_rows = ""
            for r in adjusted_rows:
                table_rows += (
                    f"<tr>"
                    f"<td style='padding: 8px; border: 1px solid #ddd;'>{r.order_number}</td>"
                    f"<td style='padding: 8px; border: 1px solid #ddd;'>{r.product_code}</td>"
                    f"<td style='padding: 8px; border: 1px solid #ddd;'>{display_phase_name(r.phase)}</td>"
                    f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>"
                    f"{r.original_planned_qty}</td>"
                    f"<td style='padding: 8px; border: 1px solid #ddd; text-align: center; "
                    f"font-weight: bold; color: #e67e22;'>{r.planned_qty_day}</td>"
                    f"<td style='padding: 8px; border: 1px solid #ddd; font-size: 12px; color: #666;'>"
                    f"Traceability shows only {r.planned_qty_day} units remaining to complete "
                    f"the order (Excel planned {r.original_planned_qty})</td>"
                    f"</tr>"
                )

            body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; color: #333;">
                <div style="margin-bottom: 20px;">
                    <img src="cid:company_logo" alt="Company Logo" width="150"/>
                </div>

                <h2 style="color: #2c3e50;">Planning Quantity Adjustment Notice</h2>

                <p>The following production quantities have been <strong>automatically adjusted</strong>
                because the Excel planning file contained quantities higher than what the traceability
                system shows as still remaining to produce.</p>

                <p>The planned daily quantities have been replaced with the actual remaining quantities
                (QtyMissing) from the traceability database to ensure accurate production tracking.</p>

                <p><strong>Source file:</strong> {excel_file}</p>
                <p><strong>Date:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

                <table style="border-collapse: collapse; width: 100%; margin: 15px 0;">
                    <thead>
                        <tr style="background-color: #2c3e50; color: white;">
                            <th style="padding: 10px; border: 1px solid #ddd;">Order</th>
                            <th style="padding: 10px; border: 1px solid #ddd;">Product</th>
                            <th style="padding: 10px; border: 1px solid #ddd;">Phase</th>
                            <th style="padding: 10px; border: 1px solid #ddd;">Original Excel Qty</th>
                            <th style="padding: 10px; border: 1px solid #ddd;">Adjusted Qty</th>
                            <th style="padding: 10px; border: 1px solid #ddd;">Reason</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows}
                    </tbody>
                </table>

                <hr style="margin: 20px 0;"/>
                <p style="color: #666; font-size: 12px;">
                    This is an automated notification from the Production Plan Monitoring System.<br/>
                    The adjustments ensure that daily targets reflect actual remaining production needs.<br/>
                    Do not reply to this email.
                </p>
            </body>
            </html>
            """

            # Prepara allegati con logo inline
            attachments = []
            if os.path.exists(LOGO_PATH):
                attachments.append(('inline', LOGO_PATH, 'company_logo'))

            send_email(
                recipients=recipients,
                subject=subject,
                body=body,
                is_html=True,
                attachments=attachments
            )

            self._adjustment_email_sent_today = True
            logger.info("Email aggiustamento qty inviata: %d righe, %d destinatari",
                        len(adjusted_rows), len(recipients))
            return True

        except Exception as e:
            logger.error("Errore invio email aggiustamento qty: %s", e)
            return False
