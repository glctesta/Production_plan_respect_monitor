import logging
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from typing import List, Optional, Tuple

from config_manager import ConfigManager
from db_connection import DatabaseConnection

logger = logging.getLogger("PlanMonitor")


@dataclass
class SnapshotRow:
    id_order: int
    id_phase: int
    qty_processed: int
    snapshot_time: datetime
    order_number: str
    product_code: str
    phase_name: str
    phase_order: int = 0


def get_db_connection() -> DatabaseConnection:
    """Crea e restituisce un oggetto DatabaseConnection."""
    cm = ConfigManager()
    return DatabaseConnection(cm)


def resolve_order(conn, order_number: str) -> Optional[Tuple[int, str]]:
    """Risolve IdOrder e ProductCode per un numero ordine."""
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT IdOrder, ProductCode
            FROM Traceability_rs.dbo.Orders o
            INNER JOIN traceability_rs.dbo.Products p ON o.idproduct = p.idproduct
            WHERE Ordernumber = ?
        """, order_number)
        row = cursor.fetchone()
        cursor.close()
        if row:
            return (row[0], row[1])
        return None
    except Exception as e:
        logger.error("Errore resolve_order per '%s': %s", order_number, e)
        return None


def resolve_phase(conn, machine_name: str) -> Optional[int]:
    """
    Risolve IdPhase per un nome macchina.
    Strategia a 3 livelli:
    1. Cerca via TraceabilityPlanning_RS (Machine -> Phase -> traceability_rs.Phases)
    2. Cerca direttamente in traceability_rs.dbo.Phases per PhaseName esatto
    3. Cerca in traceability_rs.dbo.Phases con LIKE (match parziale)
    """
    name = machine_name.strip()

    # 1. Percorso originale via TraceabilityPlanning_RS
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT p.idphase AS phaseIdIntrasa
            FROM [TraceabilityPlanning_RS].[dbo].[Machine]
            INNER JOIN [TraceabilityPlanning_RS].[dbo].[phase]
                ON Machine.PhaseId = Phase.Phaseid
            LEFT JOIN traceability_rs.dbo.phases p
                ON [TraceabilityPlanning_RS].[dbo].[phase].PhaseName COLLATE DATABASE_DEFAULT = p.PhaseName
            WHERE Machine.MachineName = ?
        """, name)
        row = cursor.fetchone()
        cursor.close()
        if row and row[0] is not None:
            return row[0]
    except Exception as e:
        logger.warning("resolve_phase via Planning fallito per '%s': %s", name, e)

    # 2. Fallback: cerca direttamente in traceability_rs.dbo.Phases per nome esatto
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT IdPhase FROM traceability_rs.dbo.Phases
            WHERE PhaseName = ?
        """, name)
        row = cursor.fetchone()
        cursor.close()
        if row and row[0] is not None:
            logger.info("resolve_phase fallback diretto OK per '%s' -> IdPhase=%d", name, row[0])
            return row[0]
    except Exception as e:
        logger.warning("resolve_phase fallback diretto fallito per '%s': %s", name, e)

    # 3. Fallback: LIKE match (es. nome macchina contiene il nome fase o viceversa)
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT IdPhase, PhaseName FROM traceability_rs.dbo.Phases
            WHERE PhaseName LIKE ? OR ? LIKE '%' + PhaseName + '%'
        """, f"%{name}%", name)
        row = cursor.fetchone()
        cursor.close()
        if row and row[0] is not None:
            logger.info("resolve_phase fallback LIKE OK per '%s' -> IdPhase=%d (PhaseName='%s')",
                        name, row[0], row[1])
            return row[0]
    except Exception as e:
        logger.warning("resolve_phase fallback LIKE fallito per '%s': %s", name, e)

    logger.error("resolve_phase: NESSUN match per macchina '%s' in nessuna strategia", name)
    return None


def insert_snapshots(conn) -> int:
    """Inserisce snapshot di produzione corrente nella tabella ShapShots."""
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO Traceability_rs.dbo.ShapShots
            SELECT  Idorder,
                    IdPhase,
                    ISNULL(Value, 0) AS QtyProcessed,
                    GETDATE() AS SnapShotTime, 0
            FROM (
                SELECT orders.idorder,
                       Orders.ordernumber,
                       Phases.IdPhase,
                       CAST(GETDATE() AS DATE) AS Period,
                       COUNT(DISTINCT Traceability_rs.dbo.BoardLabels(Scannings.IDBoard)) AS [Value],
                       Phases.Phaseorder
                FROM Traceability_rs.dbo.Scannings
                INNER JOIN Traceability_rs.dbo.OrderPhases
                    ON Scannings.IDOrderPhase = OrderPhases.IDOrderPhase
                INNER JOIN Traceability_rs.dbo.Orders
                    ON OrderPhases.IDOrder = Orders.IDOrder
                INNER JOIN Traceability_rs.dbo.Phases
                    ON OrderPhases.IDPhase = Phases.IDPhase
                INNER JOIN Traceability_rs.dbo.Products
                    ON Orders.IDProduct = Products.IDProduct
                INNER JOIN Traceability_rs.dbo.Boards
                    ON Boards.IDBoard = Scannings.IDBoard
                WHERE Scannings.ScanTimeFinish BETWEEN
                    CAST(CAST(GETDATE() AS DATE) AS DATETIME) + CAST('07:30:00' AS DATETIME) AND
                    CAST(CAST(GETDATE() + 1 AS DATE) AS DATETIME) + CAST('07:30:00' AS DATETIME)
                    AND IsPass = 1
                GROUP BY orders.idorder, Orders.ordernumber, Phases.IdPhase, phases.PhaseOrder
            ) AS AllData
        """)
        rowcount = cursor.rowcount
        cursor.close()
        logger.info("Snapshot inseriti: %d", rowcount)
        return rowcount
    except Exception as e:
        logger.error("Errore inserimento snapshot: %s", e)
        return 0


def read_unchecked_snapshots(conn) -> List[SnapshotRow]:
    """Legge gli snapshot non ancora controllati (IsChecked = 0)."""
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT s.IdOrder, s.IdPhase, s.QtyProcessed, s.SnapShotTime,
                   o.OrderNumber,
                   p2.ProductCode,
                   ph.PhaseName,
                   ISNULL(ph.PhaseOrder, 999)
            FROM traceability_rs.dbo.ShapShots s
            INNER JOIN traceability_rs.dbo.Orders o ON s.IdOrder = o.IdOrder
            INNER JOIN traceability_rs.dbo.Products p2 ON o.IDProduct = p2.IDProduct
            INNER JOIN traceability_rs.dbo.Phases ph ON s.IdPhase = ph.IdPhase
            WHERE s.IsChecked = 0
        """)
        rows = []
        for row in cursor.fetchall():
            rows.append(SnapshotRow(
                id_order=row[0],
                id_phase=row[1],
                qty_processed=row[2] or 0,
                snapshot_time=row[3],
                order_number=str(row[4]).strip(),
                product_code=str(row[5]).strip() if row[5] else "",
                phase_name=str(row[6]).strip() if row[6] else "",
                phase_order=row[7] if row[7] is not None else 999
            ))
        cursor.close()
        logger.info("Snapshot non controllati letti: %d", len(rows))
        return rows
    except Exception as e:
        logger.error("Errore lettura snapshot non controllati: %s", e)
        return []


def get_past_production(conn, id_order: int, id_phase: int, target_date: date) -> Optional[int]:
    """
    Controlla la produzione effettiva per un ordine/fase in un giorno specifico passato.
    Usa la stessa logica della query snapshot ma come SELECT, filtrata per ordine/fase e data.
    """
    try:
        cursor = conn.cursor()
        day_start = datetime.combine(target_date, datetime.strptime("07:30", "%H:%M").time())
        day_end = day_start + timedelta(days=1)
        cursor.execute("""
            SELECT COUNT(DISTINCT Traceability_rs.dbo.BoardLabels(Scannings.IDBoard)) AS Qty
            FROM Traceability_rs.dbo.Scannings
            INNER JOIN Traceability_rs.dbo.OrderPhases
                ON Scannings.IDOrderPhase = OrderPhases.IDOrderPhase
            INNER JOIN Traceability_rs.dbo.Orders
                ON OrderPhases.IDOrder = Orders.IDOrder
            INNER JOIN Traceability_rs.dbo.Phases
                ON OrderPhases.IDPhase = Phases.IDPhase
            INNER JOIN Traceability_rs.dbo.Boards
                ON Boards.IDBoard = Scannings.IDBoard
            WHERE Scannings.ScanTimeFinish BETWEEN ? AND ?
                AND IsPass = 1
                AND Orders.IdOrder = ?
                AND Phases.IdPhase = ?
        """, day_start, day_end, id_order, id_phase)
        row = cursor.fetchone()
        cursor.close()
        if row and row[0] is not None:
            return int(row[0])
        return 0
    except Exception as e:
        logger.error("Errore get_past_production per ordine=%d fase=%d data=%s: %s",
                      id_order, id_phase, target_date, e)
        return None


def mark_checked(conn, id_orders_phases: List[Tuple[int, int]]) -> None:
    """Marca come controllati gli snapshot elaborati (IsChecked = 1)."""
    if not id_orders_phases:
        return
    try:
        cursor = conn.cursor()
        for id_order, id_phase in id_orders_phases:
            cursor.execute("""
                UPDATE traceability_rs.dbo.ShapShots
                SET IsChecked = 1
                WHERE IdOrder = ? AND IdPhase = ? AND IsChecked = 0
            """, id_order, id_phase)
        cursor.close()
        logger.info("Snapshot marcati come controllati: %d coppie ordine/fase", len(id_orders_phases))
    except Exception as e:
        logger.error("Errore aggiornamento IsChecked: %s", e)
