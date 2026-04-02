import logging
from datetime import datetime, date, time, timedelta
from dataclasses import dataclass, asdict
from typing import List, Dict, Tuple, Optional

from app_config import AppConfig
from excel_parser import PlanRow, check_order_in_past_plan, check_order_in_future_plan
from db_queries import SnapshotRow, get_past_production

logger = logging.getLogger("PlanMonitor")


@dataclass
class MonitorRow:
    order_number: str
    id_order: int
    product_code: str
    phase: str
    id_phase: int
    planning_date: date
    planned_qty_day: int
    qty_done: int
    snapshot_time: Optional[datetime]
    expected_by_now: int
    projected_end_qty: int
    projected_deficit: int
    status_color: str       # "green", "yellow", "red"
    in_excel_plan: bool
    is_out_of_plan: bool
    context_star: str = ""        # "" = none, "yellow" = future, "blue" = delay
    context_note: str = ""        # Descrizione per tooltip/email

    def to_dict(self) -> dict:
        d = asdict(self)
        d["planning_date"] = self.planning_date.isoformat() if self.planning_date else None
        d["snapshot_time"] = self.snapshot_time.isoformat() if self.snapshot_time else None
        return d


def compute_projection(
    qty_done: int,
    planned_qty: int,
    snapshot_time: datetime,
    workday_start: time,
    workday_end: time,
    total_minutes: int
) -> Tuple[int, int, int]:
    """
    Calcola proiezione fine giornata.
    Ritorna (expected_by_now, projected_end_qty, projected_deficit) come interi.
    projected_deficit > 0 significa ritardo.
    """
    if planned_qty <= 0:
        return (0, qty_done, 0)

    now = snapshot_time
    today = now.date()
    day_start = datetime.combine(today, workday_start)
    day_end = datetime.combine(today, workday_end)

    # Prima dell'inizio turno
    if now <= day_start:
        return (0, 0, 0)

    # Dopo fine turno: confronto diretto
    if now >= day_end:
        deficit = planned_qty - qty_done
        return (planned_qty, qty_done, max(0, deficit))

    elapsed_minutes = (now - day_start).total_seconds() / 60.0

    # Troppo presto per proiettare (meno di 5 minuti)
    if elapsed_minutes < 5:
        return (0, 0, 0)

    fraction = elapsed_minutes / total_minutes
    expected_by_now = int(round(planned_qty * fraction))

    if elapsed_minutes > 0:
        projected_end_qty = int(round(qty_done / fraction))
    else:
        projected_end_qty = 0

    projected_deficit = planned_qty - projected_end_qty
    if projected_deficit < 0:
        projected_deficit = 0

    return (expected_by_now, projected_end_qty, projected_deficit)


def assign_color(projected_deficit: float, red_threshold: int) -> str:
    if projected_deficit <= 0:
        return "green"
    if projected_deficit > red_threshold:
        return "red"
    return "yellow"


def _enrich_out_of_plan(snap: SnapshotRow, all_plan: List[PlanRow],
                         conn, today: date) -> Tuple[str, str]:
    """
    Per un ordine fuori piano odierno, cerca contesto nei giorni precedenti e successivi.
    Ritorna (context_star, context_note).
    - "blue" + nota = ritardo da giorni precedenti
    - "yellow" + nota = pianificato nei prossimi giorni
    - "" = nessun contesto trovato
    """
    # 1. Cerca nei 2 giorni lavorativi precedenti
    past_match = check_order_in_past_plan(all_plan, snap.order_number, today, lookback_days=2)
    if past_match:
        # Verifica produzione effettiva in quel giorno
        past_qty = None
        if conn:
            past_qty = get_past_production(conn, snap.id_order, snap.id_phase, past_match.production_date)

        note = f"Planned on {past_match.production_date.strftime('%d/%m')} (qty {past_match.planned_qty})"
        if past_qty is not None:
            note += f", produced {past_qty} that day"
            if past_qty < past_match.planned_qty:
                note += " - BEHIND SCHEDULE"
        return ("blue", note)

    # 2. Cerca nei 3 giorni lavorativi successivi
    future_match = check_order_in_future_plan(all_plan, snap.order_number, today, lookahead_days=3)
    if future_match:
        note = f"Scheduled for {future_match.production_date.strftime('%d/%m')} (qty {future_match.planned_qty})"
        return ("yellow", note)

    return ("", "")


def build_dashboard_data(
    snapshots: List[SnapshotRow],
    todays_plan: List[PlanRow],
    config: AppConfig,
    all_plan: List[PlanRow] = None,
    conn=None
) -> Tuple[List[MonitorRow], List[dict]]:
    """
    Costruisce i dati della dashboard confrontando snapshot con piano.
    all_plan: piano completo (tutte le date) per cercare contesto storico/futuro.
    conn: connessione DB per verificare produzione passata.
    Ritorna (monitor_rows, mapping_errors).
    """
    if all_plan is None:
        all_plan = []

    plan_orders: set = set()
    mapping_errors: List[dict] = []

    plan_by_order: Dict[str, List[PlanRow]] = {}
    for pr in todays_plan:
        plan_orders.add(pr.order_number)
        plan_by_order.setdefault(pr.order_number, []).append(pr)

    # Aggrega snapshot per (id_order, id_phase) prendendo il piu recente
    snapshot_agg: Dict[Tuple[int, int], SnapshotRow] = {}
    for s in snapshots:
        key = (s.id_order, s.id_phase)
        if key not in snapshot_agg or s.snapshot_time > snapshot_agg[key].snapshot_time:
            snapshot_agg[key] = s

    rows: List[MonitorRow] = []
    today = date.today()

    for key, snap in snapshot_agg.items():
        order_in_plan = snap.order_number in plan_orders

        if order_in_plan:
            plan_entries = plan_by_order.get(snap.order_number, [])
            planned_qty = 0
            matched = False
            for pe in plan_entries:
                planned_qty = pe.planned_qty
                matched = True
                break

            if len(plan_entries) > 1:
                for pe in plan_entries:
                    if pe.machine_name.lower() in snap.phase_name.lower() or \
                       snap.phase_name.lower() in pe.machine_name.lower():
                        planned_qty = pe.planned_qty
                        matched = True
                        break

            if not matched:
                planned_qty = 0

            expected, projected, deficit = compute_projection(
                qty_done=snap.qty_processed,
                planned_qty=planned_qty,
                snapshot_time=snap.snapshot_time,
                workday_start=config.workday.start,
                workday_end=config.workday.end,
                total_minutes=config.workday.total_minutes
            )
            color = assign_color(deficit, config.thresholds.red_deficit)

            rows.append(MonitorRow(
                order_number=snap.order_number,
                id_order=snap.id_order,
                product_code=snap.product_code,
                phase=snap.phase_name,
                id_phase=snap.id_phase,
                planning_date=today,
                planned_qty_day=planned_qty,
                qty_done=snap.qty_processed,
                snapshot_time=snap.snapshot_time,
                expected_by_now=expected,
                projected_end_qty=projected,
                projected_deficit=deficit,
                status_color=color,
                in_excel_plan=True,
                is_out_of_plan=False
            ))
        else:
            # Ordine fuori piano Excel - cerca contesto storico/futuro
            context_star, context_note = _enrich_out_of_plan(snap, all_plan, conn, today)

            rows.append(MonitorRow(
                order_number=snap.order_number,
                id_order=snap.id_order,
                product_code=snap.product_code,
                phase=snap.phase_name,
                id_phase=snap.id_phase,
                planning_date=today,
                planned_qty_day=0,
                qty_done=snap.qty_processed,
                snapshot_time=snap.snapshot_time,
                expected_by_now=0,
                projected_end_qty=0,
                projected_deficit=0,
                status_color="red",
                in_excel_plan=False,
                is_out_of_plan=True,
                context_star=context_star,
                context_note=context_note
            ))

    # Ordinamento: fuori piano prima, poi rosso, giallo, verde
    color_order = {"red": 0, "yellow": 1, "green": 2}
    rows.sort(key=lambda r: (
        0 if r.is_out_of_plan else 1,
        color_order.get(r.status_color, 3),
        r.phase,
        r.order_number,
        -r.projected_deficit
    ))

    logger.info("Dashboard: %d righe totali, %d fuori piano",
                len(rows), sum(1 for r in rows if r.is_out_of_plan))

    return rows, mapping_errors


def compute_summary(rows: List[MonitorRow]) -> dict:
    """Calcola i contatori sintetici."""
    summary = {"green": 0, "yellow": 0, "red": 0, "out_of_plan": 0, "total": len(rows)}
    for r in rows:
        if r.is_out_of_plan:
            summary["out_of_plan"] += 1
        elif r.status_color == "green":
            summary["green"] += 1
        elif r.status_color == "yellow":
            summary["yellow"] += 1
        elif r.status_color == "red":
            summary["red"] += 1
    return summary
