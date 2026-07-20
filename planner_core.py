import json
import math
import os
import sqlite3
import traceback
import subprocess
import webbrowser
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A6, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A6, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak


APP_NAME = "Electro-Dip Production Planner"
DB_FILE = "electro_dip_planner.db"
MAX_LOOKBACK_DAYS = 365


def parse_date(value):
    if value is None or value == "":
        raise ValueError("Blank date")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Invalid date: {value}")


def parse_time(value):
    if value is None or value == "":
        raise ValueError("Blank time")
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (int, float)):
        total_seconds = round((float(value) % 1) * 86400)
        total_seconds %= 86400
        return time(total_seconds // 3600, (total_seconds % 3600) // 60)
    text = str(value).strip().replace(".", ":")
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I %p"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass
    raise ValueError(f"Invalid time: {value}")


def combine_date_time(date_value, time_value):
    return datetime.combine(parse_date(date_value), parse_time(time_value))


class Database:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path, timeout=30)
        self.conn.row_factory = sqlite3.Row
        self.create_schema()
        self.ensure_columns()


    def ensure_columns(self):
        """Add columns introduced in later versions without deleting user data."""
        def columns(table):
            return {
                row[1] for row in self.conn.execute(f"PRAGMA table_info({table})")
            }

        schedule_columns = columns("customer_schedules")
        if "customer_name" not in schedule_columns:
            self.conn.execute(
                "ALTER TABLE customer_schedules "
                "ADD COLUMN customer_name TEXT NOT NULL DEFAULT ''"
            )

        plan_columns = columns("production_plan")
        if "customer_name" not in plan_columns:
            self.conn.execute(
                "ALTER TABLE production_plan "
                "ADD COLUMN customer_name TEXT DEFAULT ''"
            )

        self.conn.commit()

    def close(self):
        self.conn.close()

    def create_schema(self):
        sql = """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS customer_schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            schedule_id TEXT NOT NULL,
            customer_name TEXT NOT NULL DEFAULT '',
            part_name TEXT NOT NULL,
            customer_qty REAL NOT NULL,
            due_datetime TEXT NOT NULL,
            priority INTEGER NOT NULL DEFAULT 999
        );

        CREATE TABLE IF NOT EXISTS stock_demand (
            part_name TEXT PRIMARY KEY,
            current_stock REAL NOT NULL DEFAULT 0,
            minimum_stock REAL NOT NULL DEFAULT 0,
            remarks TEXT
        );

        CREATE TABLE IF NOT EXISTS batch_config (
            part_name TEXT PRIMARY KEY,
            production_batch REAL NOT NULL,
            transportation_batch REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS process_bom (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_name TEXT NOT NULL,
            process_sequence INTEGER NOT NULL,
            operation_name TEXT NOT NULL,
            process_type TEXT NOT NULL,
            cycle_time_sec REAL NOT NULL DEFAULT 0,
            setup_time_min REAL NOT NULL DEFAULT 0,
            outsource_lead_hours REAL NOT NULL DEFAULT 0,
            qty_multiplier REAL NOT NULL DEFAULT 1,
            scrap_allowance REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS machine_recommendations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_name TEXT NOT NULL,
            operation_name TEXT NOT NULL,
            machine_name TEXT NOT NULL,
            preference_order INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS machine_downtime (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine_name TEXT NOT NULL,
            start_datetime TEXT NOT NULL,
            end_datetime TEXT NOT NULL,
            reason TEXT
        );

        CREATE TABLE IF NOT EXISTS shifts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shift_name TEXT NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS breaks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shift_name TEXT NOT NULL,
            break_name TEXT NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS holidays (
            holiday_date TEXT PRIMARY KEY,
            holiday_name TEXT
        );

        CREATE TABLE IF NOT EXISTS weekly_offs (
            day_number INTEGER PRIMARY KEY,
            day_name TEXT NOT NULL,
            is_off INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS production_plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id TEXT,
            schedule_id TEXT,
            customer_name TEXT DEFAULT '',
            shift_name TEXT,
            machine_name TEXT,
            operation_name TEXT,
            part_name TEXT,
            planned_qty REAL,
            start_datetime TEXT,
            end_datetime TEXT,
            process_sequence INTEGER,
            process_type TEXT,
            due_datetime TEXT,
            note TEXT
        );
        """
        self.conn.executescript(sql)
        self.conn.commit()

    def clear_master_data(self):
        tables = [
            "customer_schedules", "stock_demand", "batch_config", "process_bom",
            "machine_recommendations", "machine_downtime", "shifts", "breaks",
            "holidays", "weekly_offs", "production_plan"
        ]
        with self.conn:
            for table in tables:
                self.conn.execute(f"DELETE FROM {table}")

    def clear_plan(self):
        with self.conn:
            count = self.conn.execute("SELECT COUNT(*) FROM production_plan").fetchone()[0]
            self.conn.execute("DELETE FROM production_plan")
        return count

    def import_workbook(self, file_path):
        """
        Import valid rows and skip incomplete placeholder rows.

        Returns a report dictionary with record counts and warnings.
        Missing required worksheets remain a fatal error.
        """
        workbook = load_workbook(file_path, data_only=True)
        required = [
            "Customer_Schedules", "Stock_Demand", "Batch_Config", "Process_BOM",
            "Machine_Recommendations", "Machine_Downtime", "Shifts", "Breaks",
            "Holidays", "Weekly_Offs"
        ]
        missing = [name for name in required if name not in workbook.sheetnames]
        if missing:
            raise ValueError("Missing worksheets: " + ", ".join(missing))

        report = {
            "counts": {
                "Customer Schedules": 0,
                "Stock Records": 0,
                "Batch Configurations": 0,
                "Process BOM Operations": 0,
                "Machine Recommendations": 0,
                "Machine Downtime": 0,
                "Shifts": 0,
                "Breaks": 0,
                "Holidays": 0,
                "Weekly Off Rows": 0,
            },
            "blank_rows_skipped": 0,
            "placeholder_rows_skipped": 0,
            "operations_without_machines": 0,
            "warning_count": 0,
            "warnings": [],
        }

        def is_blank(value):
            return value is None or str(value).strip() == ""

        def row_is_blank(row):
            return all(is_blank(value) for value in row)

        def to_float(value, default=0.0):
            if is_blank(value):
                return float(default)
            try:
                return float(value)
            except (TypeError, ValueError):
                return float(default)

        def to_int(value, default=0):
            if is_blank(value):
                return int(default)
            try:
                return int(float(value))
            except (TypeError, ValueError):
                return int(default)

        def warn(sheet_name, row_number, message):
            report["warning_count"] += 1
            # Keep only a practical preview in memory and in the dialog.
            if len(report["warnings"]) < 100:
                report["warnings"].append(
                    f"{sheet_name}, row {row_number}: {message}"
                )

        # Import is atomic: a database error rolls everything back.
        with self.conn:
            self.clear_master_data()

            # ---------------- Customer Schedules ----------------
            ws = workbook["Customer_Schedules"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue

                if is_blank(row[0]) or is_blank(row[1]) or is_blank(row[2]):
                    warn(
                        ws.title, row_number,
                        "Skipped: Schedule ID, Customer Name and Part Name are required."
                    )
                    continue

                if to_float(row[3], 0) <= 0:
                    warn(
                        ws.title, row_number,
                        "Skipped: Customer Required Qty must be greater than zero."
                    )
                    continue

                if is_blank(row[4]) or is_blank(row[5]):
                    warn(
                        ws.title, row_number,
                        "Skipped: delivery date or delivery time is blank."
                    )
                    continue

                try:
                    due = combine_date_time(row[4], row[5])
                except ValueError as exc:
                    warn(ws.title, row_number, f"Skipped: {exc}")
                    continue

                self.conn.execute(
                    """INSERT INTO customer_schedules
                    (schedule_id, customer_name, part_name, customer_qty,
                     due_datetime, priority)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        str(row[1]).strip(),
                        str(row[2]).strip(),
                        to_float(row[3], 0),
                        due.isoformat(sep=" "),
                        to_int(row[6], 999),
                    )
                )
                report["counts"]["Customer Schedules"] += 1

            # ---------------- Stock Demand ----------------
            ws = workbook["Stock_Demand"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]):
                    warn(ws.title, row_number, "Skipped: Part Name is blank.")
                    continue

                self.conn.execute(
                    """INSERT OR REPLACE INTO stock_demand
                    (part_name, current_stock, minimum_stock, remarks)
                    VALUES (?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        to_float(row[1], 0),
                        to_float(row[2], 0),
                        str(row[3] or ""),
                    )
                )
                report["counts"]["Stock Records"] += 1

            # ---------------- Batch Configuration ----------------
            ws = workbook["Batch_Config"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]):
                    warn(ws.title, row_number, "Skipped: Part Name is blank.")
                    continue

                production_batch = to_float(row[1], 0)
                transportation_batch = to_float(row[2], 0)
                if production_batch <= 0 or transportation_batch <= 0:
                    warn(
                        ws.title, row_number,
                        "Skipped: production and transportation batch quantities must be greater than zero."
                    )
                    continue

                self.conn.execute(
                    """INSERT OR REPLACE INTO batch_config
                    (part_name, production_batch, transportation_batch)
                    VALUES (?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        production_batch,
                        transportation_batch,
                    )
                )
                report["counts"]["Batch Configurations"] += 1

            # ---------------- Process BOM ----------------
            ws = workbook["Process_BOM"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]) or is_blank(row[2]) or is_blank(row[3]):
                    warn(
                        ws.title, row_number,
                        "Skipped: Part Name, Operation Name and Process Type are required."
                    )
                    continue

                process_type = str(row[3]).strip().upper()
                if process_type not in ("INHOUSE", "OUTSOURCE"):
                    warn(
                        ws.title, row_number,
                        "Skipped: Process Type must be INHOUSE or OUTSOURCE."
                    )
                    continue

                sequence = to_int(row[1], 0)
                if sequence <= 0:
                    warn(
                        ws.title, row_number,
                        "Skipped: Process Sequence must be greater than zero."
                    )
                    continue

                self.conn.execute(
                    """INSERT INTO process_bom
                    (part_name, process_sequence, operation_name, process_type,
                     cycle_time_sec, setup_time_min, outsource_lead_hours,
                     qty_multiplier, scrap_allowance)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        sequence,
                        str(row[2]).strip(),
                        process_type,
                        to_float(row[4], 0),
                        to_float(row[5], 0),
                        to_float(row[6], 0),
                        max(to_float(row[7], 1), 1),
                        to_float(row[8], 0),
                    )
                )
                report["counts"]["Process BOM Operations"] += 1

            # ---------------- Machine Recommendations ----------------
            ws = workbook["Machine_Recommendations"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]) or is_blank(row[1]):
                    warn(
                        ws.title, row_number,
                        "Skipped: Part Name and Operation Name are required."
                    )
                    continue

                machines_added = 0
                for order, machine in enumerate(row[2:17], start=1):
                    if is_blank(machine):
                        continue
                    self.conn.execute(
                        """INSERT INTO machine_recommendations
                        (part_name, operation_name, machine_name, preference_order)
                        VALUES (?, ?, ?, ?)""",
                        (
                            str(row[0]).strip(),
                            str(row[1]).strip(),
                            str(machine).strip(),
                            order,
                        )
                    )
                    machines_added += 1
                    report["counts"]["Machine Recommendations"] += 1

                if machines_added == 0:
                    report["operations_without_machines"] += 1
                    warn(
                        ws.title, row_number,
                        "No machine entered in Machine 1 to Machine 15."
                    )

            # ---------------- Machine Downtime ----------------
            ws = workbook["Machine_Downtime"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue

                # A machine name without dates is a placeholder/master-list row.
                if not is_blank(row[0]) and all(is_blank(row[i]) for i in (1, 2, 3, 4)):
                    # Machine master/placeholder row, not a downtime record.
                    report["placeholder_rows_skipped"] += 1
                    continue

                if is_blank(row[0]):
                    warn(ws.title, row_number, "Skipped: Machine Name is blank.")
                    continue

                if any(is_blank(row[i]) for i in (1, 2, 3, 4)):
                    warn(
                        ws.title, row_number,
                        "Skipped: downtime start/end date and time must all be entered."
                    )
                    continue

                try:
                    start_dt = combine_date_time(row[1], row[2])
                    end_dt = combine_date_time(row[3], row[4])
                except ValueError as exc:
                    warn(ws.title, row_number, f"Skipped: {exc}")
                    continue

                if end_dt <= start_dt:
                    warn(
                        ws.title, row_number,
                        "Skipped: downtime end must be after downtime start."
                    )
                    continue

                self.conn.execute(
                    """INSERT INTO machine_downtime
                    (machine_name, start_datetime, end_datetime, reason)
                    VALUES (?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        start_dt.isoformat(sep=" "),
                        end_dt.isoformat(sep=" "),
                        str(row[5] or ""),
                    )
                )
                report["counts"]["Machine Downtime"] += 1

            # ---------------- Shifts ----------------
            ws = workbook["Shifts"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]):
                    warn(ws.title, row_number, "Skipped: Shift Name is blank.")
                    continue
                if is_blank(row[1]) or is_blank(row[2]):
                    warn(
                        ws.title, row_number,
                        "Skipped: shift start or end time is blank."
                    )
                    continue
                try:
                    start_time = parse_time(row[1]).strftime("%H:%M")
                    end_time = parse_time(row[2]).strftime("%H:%M")
                except ValueError as exc:
                    warn(ws.title, row_number, f"Skipped: {exc}")
                    continue

                self.conn.execute(
                    """INSERT INTO shifts
                    (shift_name, start_time, end_time, active)
                    VALUES (?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        start_time,
                        end_time,
                        1 if str(row[3]).strip().upper() == "Y" else 0,
                    )
                )
                report["counts"]["Shifts"] += 1

            # ---------------- Breaks ----------------
            ws = workbook["Breaks"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]) or is_blank(row[1]):
                    warn(
                        ws.title, row_number,
                        "Skipped: Shift Name and Break Name are required."
                    )
                    continue
                if is_blank(row[2]) or is_blank(row[3]):
                    warn(
                        ws.title, row_number,
                        "Skipped: break start or end time is blank."
                    )
                    continue
                try:
                    start_time = parse_time(row[2]).strftime("%H:%M")
                    end_time = parse_time(row[3]).strftime("%H:%M")
                except ValueError as exc:
                    warn(ws.title, row_number, f"Skipped: {exc}")
                    continue

                self.conn.execute(
                    """INSERT INTO breaks
                    (shift_name, break_name, start_time, end_time)
                    VALUES (?, ?, ?, ?)""",
                    (
                        str(row[0]).strip(),
                        str(row[1]).strip(),
                        start_time,
                        end_time,
                    )
                )
                report["counts"]["Breaks"] += 1

            # ---------------- Holidays ----------------
            ws = workbook["Holidays"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                if is_blank(row[0]):
                    warn(ws.title, row_number, "Skipped: Holiday Date is blank.")
                    continue
                try:
                    holiday_date = parse_date(row[0]).isoformat()
                except ValueError as exc:
                    warn(ws.title, row_number, f"Skipped: {exc}")
                    continue

                self.conn.execute(
                    """INSERT OR REPLACE INTO holidays
                    (holiday_date, holiday_name)
                    VALUES (?, ?)""",
                    (holiday_date, str(row[1] or ""))
                )
                report["counts"]["Holidays"] += 1

            # ---------------- Weekly Offs ----------------
            ws = workbook["Weekly_Offs"]
            for row_number, row in enumerate(
                ws.iter_rows(min_row=3, values_only=True), start=3
            ):
                if row_is_blank(row):
                    report["blank_rows_skipped"] += 1
                    continue
                day_number = to_int(row[0], 0)
                if day_number < 1 or day_number > 7:
                    warn(
                        ws.title, row_number,
                        "Skipped: Day Number must be between 1 and 7."
                    )
                    continue

                self.conn.execute(
                    """INSERT OR REPLACE INTO weekly_offs
                    (day_number, day_name, is_off)
                    VALUES (?, ?, ?)""",
                    (
                        day_number,
                        str(row[1] or "").strip(),
                        1 if str(row[2]).strip().upper() == "Y" else 0,
                    )
                )
                report["counts"]["Weekly Off Rows"] += 1

        return report


    def validate(self):
        errors = []
        schedule_count = self.conn.execute("SELECT COUNT(*) FROM customer_schedules").fetchone()[0]
        if schedule_count == 0:
            errors.append("No customer schedules imported.")

        active_shift_count = self.conn.execute(
            "SELECT COUNT(*) FROM shifts WHERE active = 1"
        ).fetchone()[0]
        if active_shift_count == 0:
            errors.append("No active shifts configured.")

        rows = self.conn.execute(
            """SELECT part_name, operation_name, process_type
               FROM process_bom"""
        ).fetchall()
        for row in rows:
            if row["process_type"] not in ("INHOUSE", "OUTSOURCE"):
                errors.append(
                    f"Invalid process type for {row['part_name']} / {row['operation_name']}."
                )
            if row["process_type"] == "INHOUSE":
                count = self.conn.execute(
                    """SELECT COUNT(*) FROM machine_recommendations
                       WHERE part_name = ? AND operation_name = ?""",
                    (row["part_name"], row["operation_name"])
                ).fetchone()[0]
                if count == 0:
                    errors.append(
                        f"No recommended machine for {row['part_name']} / {row['operation_name']}."
                    )
        return errors

    def generate_plan(self):
        errors = self.validate()
        if errors:
            raise ValueError("\n".join(errors))

        schedules = self.conn.execute(
            """SELECT * FROM customer_schedules
               ORDER BY due_datetime, priority"""
        ).fetchall()

        shifts = [
            (r["shift_name"], parse_time(r["start_time"]), parse_time(r["end_time"]))
            for r in self.conn.execute("SELECT * FROM shifts WHERE active = 1")
        ]
        break_map = defaultdict(list)
        for r in self.conn.execute("SELECT * FROM breaks"):
            break_map[r["shift_name"]].append(
                (parse_time(r["start_time"]), parse_time(r["end_time"]))
            )
        holidays = {
            parse_date(r["holiday_date"])
            for r in self.conn.execute("SELECT holiday_date FROM holidays")
        }
        weekly_offs = {
            r["day_number"] - 1
            for r in self.conn.execute("SELECT * FROM weekly_offs WHERE is_off = 1")
        }
        downtime_map = defaultdict(list)
        for r in self.conn.execute("SELECT * FROM machine_downtime"):
            downtime_map[r["machine_name"]].append(
                (datetime.fromisoformat(r["start_datetime"]),
                 datetime.fromisoformat(r["end_datetime"]))
            )

        reserved = defaultdict(set)

        def in_interval(current, start, end):
            if start < end:
                return start <= current < end
            return current >= start or current < end

        def get_shift(dt):
            current = dt.time()
            for shift_name, start, end in shifts:
                if in_interval(current, start, end):
                    return shift_name
            return ""

        def is_working_minute(dt):
            if dt.date() in holidays or dt.weekday() in weekly_offs:
                return False
            shift_name = get_shift(dt)
            if not shift_name:
                return False
            return not any(
                in_interval(dt.time(), start, end)
                for start, end in break_map.get(shift_name, [])
            )

        def is_downtime(machine_name, dt):
            return any(
                start <= dt < end
                for start, end in downtime_map.get(machine_name, [])
            )

        def find_slot(machine_name, end_dt, required_minutes):
            cursor = end_dt.replace(second=0, microsecond=0) - timedelta(minutes=1)
            limit = end_dt - timedelta(days=MAX_LOOKBACK_DAYS)
            minutes = []
            while cursor >= limit and len(minutes) < required_minutes:
                if (
                    is_working_minute(cursor)
                    and not is_downtime(machine_name, cursor)
                    and cursor not in reserved[machine_name]
                ):
                    minutes.append(cursor)
                cursor -= timedelta(minutes=1)
            if len(minutes) < required_minutes:
                return None
            return min(minutes), max(minutes) + timedelta(minutes=1), minutes

        with self.conn:
            self.conn.execute("DELETE FROM production_plan")

            for schedule in schedules:
                part_name = schedule["part_name"]
                stock = self.conn.execute(
                    "SELECT * FROM stock_demand WHERE part_name = ?",
                    (part_name,)
                ).fetchone()
                batch = self.conn.execute(
                    "SELECT * FROM batch_config WHERE part_name = ?",
                    (part_name,)
                ).fetchone()
                current_stock = stock["current_stock"] if stock else 0
                minimum_stock = stock["minimum_stock"] if stock else 0
                production_batch = batch["production_batch"] if batch else 0
                transportation_batch = batch["transportation_batch"] if batch else 0

                demand = max(
                    0,
                    schedule["customer_qty"] + minimum_stock - current_stock
                )
                if demand <= 0:
                    continue
                if production_batch <= 0:
                    production_batch = demand
                if transportation_batch <= 0:
                    transportation_batch = production_batch

                total_qty = math.ceil(demand / production_batch) * production_batch
                remaining = total_qty
                lot_number = 1
                due_dt = datetime.fromisoformat(schedule["due_datetime"])

                operations = self.conn.execute(
                    """SELECT * FROM process_bom
                       WHERE part_name = ?
                       ORDER BY process_sequence DESC""",
                    (part_name,)
                ).fetchall()

                while remaining > 0.000001:
                    lot_qty = min(transportation_batch, remaining)
                    operation_end = due_dt

                    for operation in operations:
                        process_qty = (
                            lot_qty
                            * max(operation["qty_multiplier"], 1)
                            * (1 + operation["scrap_allowance"] / 100)
                        )

                        if operation["process_type"] == "OUTSOURCE":
                            operation_start = operation_end - timedelta(
                                hours=operation["outsource_lead_hours"]
                            )
                            selected_end = operation_end
                            machine_name = "OUTSOURCE-" + operation["operation_name"]
                            shift_name = "OUTSOURCE"
                            note = "Outsource calendar-hour lead time"
                        else:
                            required_minutes = max(
                                1,
                                math.ceil(
                                    process_qty * operation["cycle_time_sec"] / 60
                                    + operation["setup_time_min"]
                                )
                            )
                            machines = self.conn.execute(
                                """SELECT machine_name, preference_order
                                   FROM machine_recommendations
                                   WHERE part_name = ? AND operation_name = ?
                                   ORDER BY preference_order""",
                                (part_name, operation["operation_name"])
                            ).fetchall()

                            best = None
                            for machine in machines:
                                result = find_slot(
                                    machine["machine_name"],
                                    operation_end,
                                    required_minutes
                                )
                                if result:
                                    start_dt, end_dt, minute_list = result
                                    candidate = (
                                        start_dt, end_dt, minute_list,
                                        machine["machine_name"],
                                        machine["preference_order"]
                                    )
                                    if (
                                        best is None
                                        or candidate[0] > best[0]
                                        or (
                                            candidate[0] == best[0]
                                            and candidate[4] < best[4]
                                        )
                                    ):
                                        best = candidate

                            if best is None:
                                self.conn.execute(
                                    """INSERT INTO production_plan
                                    (plan_id, schedule_id, customer_name,
                                     operation_name, part_name, planned_qty,
                                     process_sequence, process_type,
                                     due_datetime, note)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                    (
                                        "EXCEPTION",
                                        schedule["schedule_id"],
                                        schedule["customer_name"],
                                        operation["operation_name"],
                                        part_name,
                                        lot_qty,
                                        operation["process_sequence"],
                                        operation["process_type"],
                                        due_dt.isoformat(sep=" "),
                                        "No feasible machine slot."
                                    )
                                )
                                break

                            operation_start, selected_end, minute_list, machine_name, _ = best
                            reserved[machine_name].update(minute_list)
                            shift_name = get_shift(operation_start)
                            note = "Backward scheduled"

                        plan_id = (
                            f"{schedule['schedule_id']}-"
                            f"L{lot_number:03d}-"
                            f"S{operation['process_sequence']:03d}"
                        )
                        self.conn.execute(
                            """INSERT INTO production_plan
                            (plan_id, schedule_id, customer_name, shift_name,
                             machine_name, operation_name, part_name,
                             planned_qty, start_datetime, end_datetime,
                             process_sequence, process_type, due_datetime, note)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (
                                plan_id,
                                schedule["schedule_id"],
                                schedule["customer_name"],
                                shift_name,
                                machine_name,
                                operation["operation_name"],
                                part_name,
                                lot_qty,
                                operation_start.isoformat(sep=" "),
                                selected_end.isoformat(sep=" "),
                                operation["process_sequence"],
                                operation["process_type"],
                                due_dt.isoformat(sep=" "),
                                note,
                            )
                        )
                        operation_end = operation_start

                    remaining -= lot_qty
                    lot_number += 1

        return self.conn.execute(
            "SELECT COUNT(*) FROM production_plan WHERE plan_id <> 'EXCEPTION'"
        ).fetchone()[0]

    def plan_rows(self):
        return self.conn.execute(
            """SELECT * FROM production_plan
               ORDER BY COALESCE(start_datetime, '9999-12-31'), process_sequence"""
        ).fetchall()

    def todays_slips(self, selected_date):
        selected_date = parse_date(selected_date)
        rows = []
        for row in self.conn.execute(
            """SELECT p.*, COALESCE(s.priority, 999) AS priority
               FROM production_plan p
               LEFT JOIN customer_schedules s
                 ON p.schedule_id = s.schedule_id
               WHERE p.process_type = 'INHOUSE'
               ORDER BY p.machine_name, p.shift_name, p.start_datetime"""
        ):
            start = datetime.fromisoformat(row["start_datetime"])
            end = datetime.fromisoformat(row["end_datetime"])
            if start.date() <= selected_date <= end.date():
                rows.append(row)
        return rows

    def create_machine_slip_pdf(self, selected_date, output_path):
        selected_date = parse_date(selected_date)
        rows = self.todays_slips(selected_date)
        if not rows:
            raise ValueError(
                f"No in-house machine operations found for {selected_date.isoformat()}."
            )

        page_size = landscape(A6)
        doc = SimpleDocTemplate(
            output_path,
            pagesize=page_size,
            rightMargin=4 * mm,
            leftMargin=4 * mm,
            topMargin=3 * mm,
            bottomMargin=3 * mm,
        )

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "SlipTitle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=11.5,
            leading=12.5,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#FFFFFF"),
            spaceAfter=0,
        )
        subtitle_style = ParagraphStyle(
            "SlipSubtitle",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=6.8,
            leading=7.5,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1F4E78"),
        )
        label_style = ParagraphStyle(
            "SlipLabel",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=6.5,
            leading=7.2,
            textColor=colors.HexColor("#1F1F1F"),
        )
        value_style = ParagraphStyle(
            "SlipValue",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=6.5,
            leading=7.2,
            textColor=colors.HexColor("#000000"),
        )
        machine_style = ParagraphStyle(
            "MachineName",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=14,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#C00000"),
        )

        story = []
        for index, row in enumerate(rows, start=1):
            start_dt = datetime.fromisoformat(row["start_datetime"])
            end_dt = datetime.fromisoformat(row["end_datetime"])
            slip_no = f"SLIP-{selected_date:%Y%m%d}-{index:03d}"

            header = Table(
                [[Paragraph("ELECTRO-DIP", title_style)]],
                colWidths=[140 * mm],
                rowHeights=[8 * mm],
            )
            header.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1F4E78")),
                ("BOX", (0, 0), (-1, -1), 1.2, colors.HexColor("#1F4E78")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            story.append(header)

            sub_header = Table(
                [[Paragraph("MACHINE PRODUCTION SLIP", subtitle_style),
                  Paragraph(f"Slip No: {slip_no}", subtitle_style)]],
                colWidths=[78 * mm, 62 * mm],
            )
            sub_header.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#D9EAF7")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#5B9BD5")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B7B7B7")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]))
            story.append(sub_header)
            story.append(Spacer(1, 1 * mm))

            machine_box = Table(
                [[Paragraph(str(row["machine_name"]), machine_style)]],
                colWidths=[140 * mm],
                rowHeights=[9 * mm],
            )
            machine_box.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF2CC")),
                ("BOX", (0, 0), (-1, -1), 1.2, colors.HexColor("#BF9000")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            story.append(machine_box)
            story.append(Spacer(1, 1 * mm))

            detail_data = [
                [
                    Paragraph("Date", label_style),
                    Paragraph(selected_date.strftime("%d-%b-%Y"), value_style),
                    Paragraph("Shift", label_style),
                    Paragraph(str(row["shift_name"]), value_style),
                ],
                [
                    Paragraph("Part", label_style),
                    Paragraph(str(row["part_name"]), value_style),
                    Paragraph("Planned Qty", label_style),
                    Paragraph(f"{row['planned_qty']:g}", value_style),
                ],
                [
                    Paragraph("Operation", label_style),
                    Paragraph(str(row["operation_name"]), value_style),
                    Paragraph("Sequence", label_style),
                    Paragraph(str(row["process_sequence"]), value_style),
                ],
                [
                    Paragraph("Start", label_style),
                    Paragraph(start_dt.strftime("%d-%b %H:%M"), value_style),
                    Paragraph("End", label_style),
                    Paragraph(end_dt.strftime("%d-%b %H:%M"), value_style),
                ],
                [
                    Paragraph("Schedule ID", label_style),
                    Paragraph(str(row["schedule_id"]), value_style),
                    Paragraph("Priority", label_style),
                    Paragraph(
                        str(self.conn.execute(
                            "SELECT priority FROM customer_schedules WHERE schedule_id = ? LIMIT 1",
                            (row["schedule_id"],)
                        ).fetchone()[0]),
                        value_style
                    ),
                ],
            ]
            details = Table(
                detail_data,
                colWidths=[19 * mm, 51 * mm, 21 * mm, 49 * mm],
                rowHeights=[5.5 * mm] * len(detail_data),
            )
            details.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#7F7F7F")),
                ("INNERGRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#B7B7B7")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E2F0D9")),
                ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#E2F0D9")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]))
            story.append(details)
            story.append(Spacer(1, 1 * mm))

            supervisor_data = [
                [
                    Paragraph("Supervisor Instruction", label_style),
                    Paragraph("_____________________________________________", value_style),
                ],
                [
                    Paragraph("Actual Qty", label_style),
                    Paragraph("______________", value_style),
                ],
                [
                    Paragraph("Status", label_style),
                    Paragraph("Not Started / Running / Completed / Hold", value_style),
                ],
                [
                    Paragraph("Operator Sign", label_style),
                    Paragraph("____________________", value_style),
                ],
            ]
            supervisor = Table(
                supervisor_data,
                colWidths=[32 * mm, 108 * mm],
                rowHeights=[5.5 * mm, 5.5 * mm, 5.5 * mm, 5.5 * mm],
            )
            supervisor.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#7F7F7F")),
                ("INNERGRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#B7B7B7")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#FCE4D6")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ]))
            story.append(supervisor)


            if index < len(rows):
                story.append(PageBreak())

        doc.build(story)
        return len(rows)

    def export_plan(self, output_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Production_Plan"
        headers = [
            "Plan ID", "Schedule ID", "Customer Name", "Shift Name",
            "Machine Name", "Operation Name", "Part Name", "Planned Qty",
            "Production Start", "Production End", "Process Sequence",
            "Process Type", "Due Date/Time", "Status / Note"
        ]
        ws.append(headers)
        for row in self.plan_rows():
            ws.append([
                row["plan_id"], row["schedule_id"], row["customer_name"],
                row["shift_name"], row["machine_name"],
                row["operation_name"], row["part_name"],
                row["planned_qty"], row["start_datetime"], row["end_datetime"],
                row["process_sequence"], row["process_type"],
                row["due_datetime"], row["note"]
            ])

        slip_ws = wb.create_sheet("Machine_Slips")
        slip_ws.append([
            "Machine", "Shift", "Customer", "Operation", "Part",
            "Plan Qty", "Actual Qty", "Rejected Qty", "Start", "End",
            "Schedule ID", "Priority", "Status", "Remarks"
        ])
        today = date.today()
        for row in self.todays_slips(today):
            slip_ws.append([
                row["machine_name"], row["shift_name"], row["customer_name"],
                row["operation_name"], row["part_name"], row["planned_qty"],
                "", "", row["start_datetime"], row["end_datetime"],
                row["schedule_id"], row["priority"], "", ""
            ])

        for sheet in wb.worksheets:
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="5B9BD5")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            for column in range(1, sheet.max_column + 1):
                max_length = max(
                    len(str(sheet.cell(row, column).value or ""))
                    for row in range(1, sheet.max_row + 1)
                )
                sheet.column_dimensions[get_column_letter(column)].width = min(
                    max(max_length + 2, 12), 35
                )
        wb.save(output_path)


def fit_text(canvas_obj, text, max_width, font_name="Helvetica-Bold",
             max_size=11, min_size=6):
    text = str(text or "")
    size = max_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    canvas_obj.setFont(font_name, size)
    return size


def draw_field(c, x, y, width, height, label, value,
               label_fill=colors.HexColor("#D9EAF7"),
               value_fill=colors.white,
               border=colors.HexColor("#7F8C8D"),
               value_font="Helvetica-Bold"):
    c.setStrokeColor(border)
    c.setLineWidth(0.55)
    c.setFillColor(label_fill)
    c.rect(x, y, width * 0.35, height, stroke=1, fill=1)
    c.setFillColor(value_fill)
    c.rect(x + width * 0.35, y, width * 0.65, height, stroke=1, fill=1)

    c.setFillColor(colors.HexColor("#1F1F1F"))
    c.setFont("Helvetica-Bold", 6.5)
    c.drawString(x + 2.2 * mm, y + height / 2 - 2.1, str(label))

    value_text = str(value or "")
    fit_text(c, value_text, width * 0.62 - 4 * mm,
             font_name=value_font, max_size=8.5, min_size=5.5)
    c.drawString(x + width * 0.35 + 2.2 * mm,
                 y + height / 2 - 2.4, value_text)


def draw_checkbox(c, x, y, label, checked=False):
    box = 4 * mm
    c.setStrokeColor(colors.HexColor("#34495E"))
    c.setFillColor(colors.white)
    c.rect(x, y, box, box, stroke=1, fill=1)
    if checked:
        c.setStrokeColor(colors.HexColor("#1E8449"))
        c.setLineWidth(1.3)
        c.line(x + 0.8 * mm, y + 2 * mm, x + 1.8 * mm, y + 0.8 * mm)
        c.line(x + 1.8 * mm, y + 0.8 * mm, x + 3.4 * mm, y + 3.3 * mm)
    c.setFillColor(colors.HexColor("#2C3E50"))
    c.setFont("Helvetica", 6.5)
    c.drawString(x + box + 1.4 * mm, y + 1.1 * mm, label)


def create_machine_slips_pdf(rows, selected_date, output_path):
    """
    Print only the approved ELECTRO-DIP grouped machine slip format.

    Grouping:
      selected date + machine + shift = one A4 landscape page.
    """
    from collections import defaultdict

    grouped = defaultdict(list)
    for row in rows:
        grouped[(row["machine_name"], row["shift_name"])].append(row)

    page_width, page_height = landscape((297 * mm, 210 * mm))
    c = canvas.Canvas(str(output_path), pagesize=(page_width, page_height))
    c.setTitle(f"ELECTRO-DIP Machine Slips {selected_date.isoformat()}")

    logo_path = Path(__file__).resolve().parent / "electro_dip_logo.png"

    def cell(x, y, w, h, text="", fill=colors.white,
             font="Helvetica", size=6.4, align="center",
             text_color=colors.HexColor("#1F1F1F"),
             border=colors.HexColor("#7F8C8D")):
        c.setStrokeColor(border)
        c.setLineWidth(0.45)
        c.setFillColor(fill)
        c.rect(x, y, w, h, stroke=1, fill=1)

        c.setFillColor(text_color)
        fit_text(c, text, w - 3 * mm, font_name=font,
                 max_size=size, min_size=4.2)
        baseline = y + h / 2 - 2.0
        if align == "left":
            c.drawString(x + 1.5 * mm, baseline, str(text or ""))
        elif align == "right":
            c.drawRightString(x + w - 1.5 * mm, baseline, str(text or ""))
        else:
            c.drawCentredString(x + w / 2, baseline, str(text or ""))

    def field(x, y, w, h, label, value, label_fill):
        label_w = w * 0.36
        cell(x, y, label_w, h, label, label_fill,
             font="Helvetica-Bold", size=6.2, align="left")
        cell(x + label_w, y, w - label_w, h, value, colors.white,
             font="Helvetica-Bold", size=8.2, align="left")

    for slip_no, ((machine_name, shift_name), operations) in enumerate(
        sorted(grouped.items()), start=1
    ):
        operations.sort(key=lambda row: row["start_datetime"])

        # Page background and border.
        c.setFillColor(colors.HexColor("#F7F9FB"))
        c.rect(0, 0, page_width, page_height, stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor("#1F4E78"))
        c.setLineWidth(1.2)
        c.roundRect(5 * mm, 5 * mm, page_width - 10 * mm,
                    page_height - 10 * mm, 3 * mm, stroke=1, fill=0)

        # Blue header exactly as approved.
        header_y = page_height - 30 * mm
        c.setFillColor(colors.HexColor("#1F4E78"))
        c.roundRect(5 * mm, header_y, page_width - 10 * mm,
                    25 * mm, 3 * mm, stroke=0, fill=1)

        if logo_path.exists():
            c.setFillColor(colors.white)
            c.rect(9 * mm, header_y + 4 * mm, 18 * mm, 17 * mm,
                   stroke=0, fill=1)
            c.drawImage(str(logo_path), 10 * mm, header_y + 5 * mm,
                        width=16 * mm, height=15 * mm,
                        preserveAspectRatio=True, mask="auto")

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(34 * mm, header_y + 15 * mm, "ELECTRO-DIP")
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(34 * mm, header_y + 8 * mm,
                     "MACHINE DAILY PRODUCTION SLIP")

        c.setFont("Helvetica-Bold", 7)
        c.drawRightString(
            page_width - 10 * mm, header_y + 16 * mm,
            f"Slip No.: MS-{selected_date:%Y%m%d}-{slip_no:03d}"
        )
        c.drawRightString(
            page_width - 10 * mm, header_y + 10 * mm,
            f"Date: {selected_date:%d-%b-%Y}"
        )

        # Header fields: Customer intentionally excluded.
        info_y = header_y - 12 * mm
        field(8 * mm, info_y, 92 * mm, 9 * mm,
              "Machine", machine_name, colors.HexColor("#D9EAF7"))
        field(103 * mm, info_y, 53 * mm, 9 * mm,
              "Shift", shift_name, colors.HexColor("#E4DFEC"))
        field(159 * mm, info_y, 62 * mm, 9 * mm,
              "Operator", "", colors.HexColor("#E2F0D9"))
        field(224 * mm, info_y, 65 * mm, 9 * mm,
              "Supervisor", "", colors.HexColor("#FFF2CC"))

        # Operations table.
        table_top = info_y - 6 * mm
        table_x = 8 * mm
        row_h = 8.4 * mm
        widths_mm = [8, 16, 16, 30, 24, 42, 18, 20, 22, 18, 18, 36]
        widths = [value * mm for value in widths_mm]
        headings = [
            "Sr", "Start", "End", "Customer", "Part No.",
            "Operation", "Plan Qty", "Actual Qty", "Rejected Qty",
            "Priority", "Status", "Remarks"
        ]

        header_row_y = table_top - row_h
        current_x = table_x
        for heading, width in zip(headings, widths):
            cell(current_x, header_row_y, width, row_h, heading,
                 colors.HexColor("#1F4E78"),
                 font="Helvetica-Bold", size=6.2,
                 text_color=colors.white)
            current_x += width

        # Same page supports 12 lines, like a practical daily slip.
        max_rows = 12
        shown_operations = operations[:max_rows]
        total_planned = 0.0

        for index, row in enumerate(shown_operations, start=1):
            y = header_row_y - index * row_h
            alternate = (
                colors.white if index % 2
                else colors.HexColor("#F2F5F7")
            )
            start_dt = datetime.fromisoformat(row["start_datetime"])
            end_dt = datetime.fromisoformat(row["end_datetime"])

            values = [
                index,
                start_dt.strftime("%H:%M"),
                end_dt.strftime("%H:%M"),
                row["customer_name"],
                row["part_name"],
                row["operation_name"],
                f"{float(row['planned_qty'] or 0):g}",
                "",
                "",
                str(row["priority"] if "priority" in row.keys() else ""),
                "",
                "",
            ]
            total_planned += float(row["planned_qty"] or 0)

            current_x = table_x
            for col_index, (value, width) in enumerate(zip(values, widths)):
                fill = alternate
                if col_index == 7:
                    fill = colors.HexColor("#E2F0D9")
                elif col_index == 8:
                    fill = colors.HexColor("#FCE4D6")
                elif col_index == 9:
                    fill = colors.HexColor("#FFF2CC")

                align = "left" if col_index in (3, 5, 11) else "center"
                font = (
                    "Helvetica-Bold"
                    if col_index in (3, 4, 5, 6)
                    else "Helvetica"
                )
                cell(current_x, y, width, row_h, value, fill,
                     font=font, size=6.2, align=align)
                current_x += width

        # Totals row.
        totals_y = header_row_y - (max(len(shown_operations), 1) + 1) * row_h - 4 * mm
        field(8 * mm, totals_y, 66 * mm, 9 * mm,
              "Total Planned Qty", f"{total_planned:g}",
              colors.HexColor("#D9EAF7"))
        field(77 * mm, totals_y, 66 * mm, 9 * mm,
              "Total Actual Qty", "",
              colors.HexColor("#E2F0D9"))
        field(146 * mm, totals_y, 66 * mm, 9 * mm,
              "Total Rejected Qty", "",
              colors.HexColor("#F4CCCC"))
        field(215 * mm, totals_y, 74 * mm, 9 * mm,
              "Machine Utilization", "____ %",
              colors.HexColor("#FFF2CC"))

        # Remarks and overall status.
        remarks_y = totals_y - 18 * mm
        remarks_w = 195 * mm
        cell(8 * mm, remarks_y, remarks_w, 15 * mm, "",
             colors.white, align="left")
        c.setFillColor(colors.HexColor("#1F4E78"))
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(10 * mm, remarks_y + 11 * mm,
                     "SUPERVISOR / QUALITY REMARKS")
        c.setStrokeColor(colors.HexColor("#BDC3C7"))
        c.line(10 * mm, remarks_y + 7 * mm,
               8 * mm + remarks_w - 3 * mm, remarks_y + 7 * mm)
        c.line(10 * mm, remarks_y + 3.5 * mm,
               8 * mm + remarks_w - 3 * mm, remarks_y + 3.5 * mm)

        status_x = 206 * mm
        status_w = 83 * mm
        cell(status_x, remarks_y, status_w, 15 * mm, "",
             colors.white)
        c.setFillColor(colors.HexColor("#2C3E50"))
        c.setFont("Helvetica", 6.2)
        c.drawString(status_x + 3 * mm, remarks_y + 10.5 * mm,
                     "Status:")
        c.drawString(status_x + 20 * mm, remarks_y + 10.5 * mm,
                     "[ ] Completed")
        c.drawString(status_x + 55 * mm, remarks_y + 10.5 * mm,
                     "[ ] Hold")
        c.drawString(status_x + 20 * mm, remarks_y + 4.5 * mm,
                     "[ ] Running")
        c.drawString(status_x + 55 * mm, remarks_y + 4.5 * mm,
                     "[ ] Rework")

        # Signatures at bottom.
        signature_y = 9 * mm
        c.setFillColor(colors.HexColor("#2C3E50"))
        c.setFont("Helvetica", 6.4)
        c.drawString(
            10 * mm, signature_y,
            "Operator Sign: ______________________________"
        )
        c.drawCentredString(
            page_width / 2, signature_y,
            "Quality Sign: ______________________________"
        )
        c.drawRightString(
            page_width - 10 * mm, signature_y,
            "Supervisor Sign: ______________________________"
        )

        if len(operations) > max_rows:
            c.setFillColor(colors.HexColor("#C0392B"))
            c.setFont("Helvetica-Bold", 6)
            c.drawRightString(
                page_width - 10 * mm, remarks_y - 4 * mm,
                f"Additional operations: {len(operations) - max_rows}"
            )

        c.showPage()

    c.save()
    return output_path

