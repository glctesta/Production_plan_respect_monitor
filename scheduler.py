import logging
import threading
import os
from datetime import datetime, date, time, timedelta
from typing import Optional, List

from app_config import AppConfig
from excel_parser import find_latest_excel, parse_last_phase, get_todays_plan, PlanRow
from db_queries import get_db_connection, insert_snapshots, read_unchecked_snapshots, mark_checked, resolve_phase, resolve_order, get_qty_missing
from monitor_engine import build_dashboard_data, compute_summary, MonitorRow
from email_alerter import EmailAlertManager

logger = logging.getLogger("PlanMonitor")

# Intervallo di verifica aggiornamento file Excel (minuti)
EXCEL_CHECK_INTERVAL_MINUTES = 30


class CycleOrchestrator:
    def __init__(self, config: AppConfig, email_alerter: EmailAlertManager):
        self.config = config
        self.email_alerter = email_alerter
        self._lock = threading.Lock()
        self._cycle_running = False

        # Cache Excel
        self._cached_excel_path: Optional[str] = None
        self._cached_excel_mtime: Optional[float] = None  # timestamp os.path.getmtime
        self._cached_all_plan: List[PlanRow] = []
        self._last_excel_check: Optional[datetime] = None

        # Stato condiviso per la dashboard (thread-safe via _lock)
        self.dashboard_data = {
            "rows": [],
            "summary": {"green": 0, "yellow": 0, "red": 0, "out_of_plan": 0, "total": 0},
            "excel_file": None,
            "excel_modified": None,
            "last_update": None,
            "cycle_running": False,
            "mapping_errors": [],
            "last_error": None
        }

    def _is_within_workday(self) -> bool:
        now = datetime.now().time()
        return self.config.workday.start <= now <= self.config.workday.end

    def _should_recheck_excel(self) -> bool:
        """Ritorna True se sono passati >= 30 min dall'ultimo controllo Excel."""
        if self._last_excel_check is None:
            return True
        elapsed = (datetime.now() - self._last_excel_check).total_seconds() / 60.0
        return elapsed >= EXCEL_CHECK_INTERVAL_MINUTES

    def _load_excel_plan(self) -> tuple:
        """
        Carica il piano Excel, usando la cache se il file non e' cambiato.
        Controlla il file solo ogni 30 minuti.
        Ritorna (excel_file_name, excel_modified_iso, all_plan) oppure (None, None, []) se errore.
        """
        if not self._should_recheck_excel() and self._cached_all_plan:
            # Usa la cache
            excel_file = os.path.basename(self._cached_excel_path) if self._cached_excel_path else None
            excel_mod_iso = datetime.fromtimestamp(self._cached_excel_mtime).isoformat() if self._cached_excel_mtime else None
            logger.info("Excel cache valida (prossimo check tra %.0f min), uso: %s",
                        EXCEL_CHECK_INTERVAL_MINUTES - (datetime.now() - self._last_excel_check).total_seconds() / 60.0,
                        excel_file)
            return (excel_file, excel_mod_iso, self._cached_all_plan)

        # Tempo di verificare se c'e' un file nuovo
        self._last_excel_check = datetime.now()
        logger.info("Verifica file Excel in %s...", self.config.planning.folder)

        result = find_latest_excel(self.config.planning.folder)
        if result is None:
            return (None, None, [])

        excel_path, excel_mod_dt = result
        excel_file = os.path.basename(excel_path)

        try:
            current_mtime = os.path.getmtime(excel_path)
        except OSError:
            current_mtime = 0.0

        # Confronta con la cache: stesso file e stessa data modifica?
        if (self._cached_excel_path == excel_path
                and self._cached_excel_mtime is not None
                and abs(current_mtime - self._cached_excel_mtime) < 1.0
                and self._cached_all_plan):
            logger.info("File Excel invariato: %s - uso cache (%d righe piano)",
                        excel_file, len(self._cached_all_plan))
            return (excel_file, excel_mod_dt.isoformat(), self._cached_all_plan)

        # File nuovo o modificato: ri-parsa
        if self._cached_excel_path and self._cached_excel_path != excel_path:
            logger.info("NUOVO file Excel rilevato: %s (precedente: %s)",
                        excel_file, os.path.basename(self._cached_excel_path))
        elif self._cached_excel_mtime and abs(current_mtime - self._cached_excel_mtime) >= 1.0:
            logger.info("File Excel AGGIORNATO: %s (mtime cambiato)", excel_file)
        else:
            logger.info("Prima lettura file Excel: %s", excel_file)

        all_plan = parse_last_phase(excel_path, self.config.planning.sheet)

        # Aggiorna cache
        self._cached_excel_path = excel_path
        self._cached_excel_mtime = current_mtime
        self._cached_all_plan = all_plan

        logger.info("Excel parsato: %d righe piano totali", len(all_plan))
        return (excel_file, excel_mod_dt.isoformat(), all_plan)

    def run_cycle(self, force: bool = False) -> dict:
        """
        Esegue un ciclo completo di monitoraggio.
        force=True bypassa il controllo orario (per il bottone Esegui Ora).
        """
        if not force and not self._is_within_workday():
            logger.info("Fuori fascia oraria lavorativa (%s-%s), ciclo saltato",
                        self.config.workday.start, self.config.workday.end)
            return self.dashboard_data

        if self._cycle_running:
            logger.warning("Ciclo gia in esecuzione, richiesta ignorata")
            return self.dashboard_data

        with self._lock:
            self._cycle_running = True
            self.dashboard_data["cycle_running"] = True

        cycle_start = datetime.now()
        logger.info("=== INIZIO CICLO === %s", cycle_start.strftime("%Y-%m-%d %H:%M:%S"))

        try:
            # 1. Carica piano Excel (con cache 30 min)
            excel_file, excel_modified, all_plan = self._load_excel_plan()
            if excel_file is None:
                error_msg = "Nessun file Excel trovato in " + self.config.planning.folder
                logger.error(error_msg)
                with self._lock:
                    self.dashboard_data["last_error"] = error_msg
                    self.dashboard_data["last_update"] = datetime.now().isoformat()
                    self._cycle_running = False
                    self.dashboard_data["cycle_running"] = False
                return self.dashboard_data

            if not all_plan:
                logger.warning("Nessuna riga piano letta dal file Excel")

            # 2. Filtra piano odierno
            todays_plan = get_todays_plan(all_plan)
            logger.info("Piano odierno: %d righe", len(todays_plan))

            # 3. Operazioni DB
            db_conn = get_db_connection()
            try:
                conn = db_conn.connect()

                # 4. Risolvi fasi Excel -> id_phase dal DB
                #    Costruisce mappa (order_number, id_phase) -> PlanRow
                resolved_plan = {}  # {(order_number, id_phase): PlanRow}
                phase_cache = {}    # {machine_name: id_phase}
                unresolved_count = 0
                for pr in todays_plan:
                    if pr.machine_name not in phase_cache:
                        resolved_id = resolve_phase(conn, pr.machine_name)
                        phase_cache[pr.machine_name] = resolved_id
                        logger.info("  resolve_phase('%s') -> %s", pr.machine_name, resolved_id)
                    id_phase = phase_cache[pr.machine_name]
                    if id_phase is not None:
                        resolved_plan[(pr.order_number, id_phase)] = pr
                    else:
                        unresolved_count += 1
                        logger.warning("Fase non risolta per macchina '%s' (ordine %s)",
                                       pr.machine_name, pr.order_number)

                logger.info("Piano odierno risolto: %d coppie ordine/fase, %d non risolte (su %d righe Excel)",
                            len(resolved_plan), unresolved_count, len(todays_plan))

                if len(resolved_plan) > 0:
                    for rk in list(resolved_plan.keys())[:10]:
                        logger.info("  resolved_plan key: order='%s' id_phase=%d", rk[0], rk[1])

                if len(resolved_plan) == 0 and len(todays_plan) > 0:
                    unique_machines = set(pr.machine_name for pr in todays_plan)
                    logger.error("NESSUNA fase risolta! Macchine nel piano Excel: %s", unique_machines)

                # 4b. Calcola QtyMissing dalla tracciabilita' per ogni ordine
                qty_missing_map = {}  # {(order_number, id_phase): qty_missing}
                order_ids_cache = {}  # {order_number: id_order}
                orders_loaded = set()
                for (order_number, id_phase), pr in resolved_plan.items():
                    if order_number not in order_ids_cache:
                        result = resolve_order(conn, order_number)
                        order_ids_cache[order_number] = result[0] if result else None
                    id_order = order_ids_cache.get(order_number)
                    if id_order is not None and id_order not in orders_loaded:
                        orders_loaded.add(id_order)
                        phase_missing = get_qty_missing(conn, id_order)
                        for ph_id, qty_m in phase_missing.items():
                            qty_missing_map[(order_number, ph_id)] = qty_m

                logger.info("QtyMissing calcolate per %d coppie ordine/fase", len(qty_missing_map))

                # 5. Inserisci snapshot
                try:
                    snap_count = insert_snapshots(conn)
                    logger.info("Snapshot inseriti: %d", snap_count)
                except Exception as e:
                    logger.error("Errore inserimento snapshot: %s", e)

                # 6. Leggi snapshot non controllati
                snapshots = read_unchecked_snapshots(conn)

                # 7. Costruisci dati dashboard (usa resolved_plan per filtrare solo fasi Excel)
                rows, mapping_errors = build_dashboard_data(
                    snapshots, todays_plan, self.config,
                    all_plan=all_plan, conn=conn,
                    resolved_plan=resolved_plan,
                    qty_missing_map=qty_missing_map
                )
                summary = compute_summary(rows)

                logger.info("Risultati: verde=%d, giallo=%d, rosso=%d, fuori_piano=%d",
                            summary["green"], summary["yellow"], summary["red"], summary["out_of_plan"])

                # 7. Aggiorna stato condiviso
                with self._lock:
                    self.dashboard_data["rows"] = [r.to_dict() for r in rows]
                    self.dashboard_data["summary"] = summary
                    self.dashboard_data["excel_file"] = excel_file
                    self.dashboard_data["excel_modified"] = excel_modified
                    self.dashboard_data["last_update"] = datetime.now().isoformat()
                    self.dashboard_data["mapping_errors"] = mapping_errors
                    self.dashboard_data["last_error"] = None

                # 8. Invia email se necessario
                try:
                    if self.config.email.enabled:
                        self.email_alerter.send_alerts(conn, rows, summary, excel_file)
                except Exception as e:
                    logger.error("Errore invio email alert: %s", e)

                # 9. Marca snapshot come controllati
                try:
                    ids_to_mark = list(set((s.id_order, s.id_phase) for s in snapshots))
                    mark_checked(conn, ids_to_mark)
                except Exception as e:
                    logger.error("Errore marcatura IsChecked: %s", e)

            finally:
                db_conn.disconnect()

        except Exception as e:
            logger.error("Errore imprevisto nel ciclo: %s", e, exc_info=True)
            with self._lock:
                self.dashboard_data["last_error"] = str(e)
                self.dashboard_data["last_update"] = datetime.now().isoformat()

        finally:
            with self._lock:
                self._cycle_running = False
                self.dashboard_data["cycle_running"] = False

            elapsed = (datetime.now() - cycle_start).total_seconds()
            logger.info("=== FINE CICLO === Durata: %.1f secondi", elapsed)

        return self.dashboard_data

    def get_status(self) -> dict:
        with self._lock:
            return dict(self.dashboard_data)
