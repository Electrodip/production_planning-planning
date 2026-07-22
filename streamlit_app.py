import io
import os
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from planner_core import Database, create_machine_slips_pdf, parse_date

APP_TITLE = "Electro-Dip Online Production Planner"
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Electro_Dip_Import_Template.xlsx"
LOGO_PATH = BASE_DIR / "electro_dip_logo.png"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)


def check_password() -> bool:
    """Optional password protection using Streamlit secrets."""
    try:
        expected = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        expected = ""
    if not expected:
        return True

    if st.session_state.get("authenticated"):
        return True

    st.title("🔒 Electro-Dip Production Planner")
    password = st.text_input("Enter app password", type="password")
    if st.button("Sign in", type="primary"):
        if password == expected:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


def get_database(session_id: str) -> Database:
    """Open a fresh SQLite connection for each Streamlit rerun.

    Streamlit may execute reruns on different threads. A cached SQLite
    connection can therefore raise sqlite3.ProgrammingError. The database
    file remains session-specific, while the connection is created in the
    current execution thread.
    """
    db_path = Path(tempfile.gettempdir()) / f"electro_dip_{session_id}.db"
    return Database(db_path)


if "session_id" not in st.session_state:
    st.session_state.session_id = os.urandom(12).hex()
if "import_report" not in st.session_state:
    st.session_state.import_report = None
if "selected_slip_date" not in st.session_state:
    st.session_state.selected_slip_date = date.today()

db = get_database(st.session_state.session_id)


def rows_to_dataframe(rows) -> pd.DataFrame:
    records = [dict(row) for row in rows]
    return pd.DataFrame(records)


def show_import_report(report: dict) -> None:
    counts = report.get("counts", {})
    count_df = pd.DataFrame(
        [{"Data Group": name, "Imported Records": value} for name, value in counts.items()]
    )
    st.success("Excel data imported successfully.")
    st.dataframe(count_df, hide_index=True, use_container_width=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Blank rows skipped", report.get("blank_rows_skipped", 0))
    c2.metric("Placeholder rows", report.get("placeholder_rows_skipped", 0))
    c3.metric("Operations without machines", report.get("operations_without_machines", 0))
    c4.metric("Warnings", report.get("warning_count", 0))

    warnings = report.get("warnings", [])
    if warnings:
        with st.expander("Import warnings (preview)", expanded=False):
            for warning in warnings[:100]:
                st.write("•", warning)
            if report.get("warning_count", len(warnings)) > len(warnings):
                st.info(
                    f"Only the first {len(warnings)} warnings are displayed. "
                    f"Total warnings: {report.get('warning_count')}"
                )


def export_plan_bytes() -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = tmp.name
    try:
        db.export_plan(path)
        return Path(path).read_bytes()
    finally:
        Path(path).unlink(missing_ok=True)


def slip_pdf_bytes(selected_date: date) -> bytes:
    rows = db.todays_slips(selected_date)
    if not rows:
        return b""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        path = Path(tmp.name)
    try:
        create_machine_slips_pdf(rows, selected_date, path)
        return path.read_bytes()
    finally:
        path.unlink(missing_ok=True)


# Header
h1, h2 = st.columns([1, 8])
with h1:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=90)
with h2:
    st.title("ELECTRO-DIP")
    st.subheader("Online Production Planning System")

st.caption(
    "Excel import • 15 recommended machines • backward planning • machine-wise approved slips"
)

# Sidebar
with st.sidebar:
    st.header("Data Import")
    uploaded = st.file_uploader(
        "Upload planning template",
        type=["xlsx", "xlsm"],
        help="Use the supplied Electro-Dip import template.",
    )

    if TEMPLATE_PATH.exists():
        st.download_button(
            "Download Import Template",
            data=TEMPLATE_PATH.read_bytes(),
            file_name="Electro_Dip_Import_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if st.button("Import Excel", type="primary", use_container_width=True):
        if uploaded is None:
            st.error("Select an Excel file first.")
        else:
            suffix = Path(uploaded.name).suffix or ".xlsx"
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp.write(uploaded.getbuffer())
                upload_path = Path(tmp.name)
            try:
                with st.spinner("Importing and checking data..."):
                    report = db.import_workbook(upload_path)
                st.session_state.import_report = report
                st.success("Import completed.")
                st.rerun()
            except Exception as exc:
                st.error(f"Import failed: {exc}")
            finally:
                upload_path.unlink(missing_ok=True)

    st.divider()
    st.header("Planning Actions")
    if st.button("Validate Inputs", use_container_width=True):
        issues = db.validate()
        if issues:
            st.session_state.validation_issues = issues
        else:
            st.session_state.validation_issues = []
            st.success("All inputs are valid.")

    if st.button("Generate Production Plan", type="primary", use_container_width=True):
        issues = db.validate()
        if issues:
            st.session_state.validation_issues = issues
            st.error(f"Planning blocked: {len(issues)} validation issue(s).")
        else:
            try:
                with st.spinner("Generating backward production plan..."):
                    count = db.generate_plan()
                st.session_state.plan_message = f"{count} operation rows created."
                st.success(st.session_state.plan_message)
                st.rerun()
            except Exception as exc:
                st.error(f"Planning failed: {exc}")

    if st.button("Clear Production Plan", use_container_width=True):
        count = db.clear_plan()
        st.success(f"{count} production-plan rows cleared successfully.")
        st.rerun()

# Metrics
schedule_count = db.conn.execute("SELECT COUNT(*) FROM customer_schedules").fetchone()[0]
part_count = db.conn.execute("SELECT COUNT(DISTINCT part_name) FROM process_bom").fetchone()[0]
operation_count = db.conn.execute("SELECT COUNT(*) FROM process_bom").fetchone()[0]
machine_count = db.conn.execute("SELECT COUNT(DISTINCT machine_name) FROM machine_recommendations").fetchone()[0]
plan_count = db.conn.execute("SELECT COUNT(*) FROM production_plan WHERE plan_id <> 'EXCEPTION'").fetchone()[0]
exception_count = db.conn.execute("SELECT COUNT(*) FROM production_plan WHERE plan_id = 'EXCEPTION'").fetchone()[0]
produced_good_qty = db.conn.execute(
    "SELECT COALESCE(SUM(good_qty), 0) FROM production_updates"
).fetchone()[0]

metrics = st.columns(7)
for column, label, value in zip(
    metrics,
    ["Schedules", "Parts", "Operations", "Machines", "Plan Rows", "Exceptions", "Good Qty Reported"],
    [schedule_count, part_count, operation_count, machine_count, plan_count, exception_count, produced_good_qty],
):
    column.metric(label, value)

if st.session_state.import_report:
    with st.expander("Latest import report", expanded=False):
        show_import_report(st.session_state.import_report)

validation_issues = st.session_state.get("validation_issues")
if validation_issues:
    st.error(f"Validation found {len(validation_issues)} issue(s).")
    with st.expander("Validation issues", expanded=True):
        for issue in validation_issues[:200]:
            st.write("•", issue)
        if len(validation_issues) > 200:
            st.info(f"Additional issues not displayed: {len(validation_issues) - 200}")

# Tabs
dashboard_tab, plan_tab, slips_tab, progress_tab, data_tab = st.tabs(
    ["Dashboard", "Production Plan", "Machine Slips", "Production Progress", "Imported Data"]
)

with dashboard_tab:
    st.subheader("Planning Overview")

    revised_df = rows_to_dataframe(db.revised_quantity_rows())
    if not revised_df.empty:
        st.markdown("#### Revised Production Quantities")
        display_columns = [
            "schedule_id", "customer_name", "part_name",
            "customer_required_qty", "minimum_stock", "current_stock",
            "original_net_requirement", "accepted_produced_qty",
            "revised_plan_qty", "due_datetime", "priority",
        ]
        revised_columns = [
            column for column in display_columns
            if column in revised_df.columns
        ]
        st.dataframe(
            revised_df[revised_columns],
            hide_index=True,
            use_container_width=True,
        )
        st.caption(
            "Revised Plan Qty = Customer Required Qty + Minimum Stock "
            "− Current Stock − Accepted Produced Qty. "
            "The final transportation lot may be smaller than the standard batch."
        )

    if plan_count:
        summary_df = pd.read_sql_query(
            """
            SELECT machine_name AS Machine,
                   shift_name AS Shift,
                   COUNT(*) AS Operations,
                   ROUND(SUM(planned_qty), 2) AS Planned_Qty,
                   MIN(start_datetime) AS First_Start,
                   MAX(end_datetime) AS Last_End
            FROM production_plan
            WHERE plan_id <> 'EXCEPTION' AND process_type='INHOUSE'
            GROUP BY machine_name, shift_name
            ORDER BY Machine, Shift
            """,
            db.conn,
        )
        st.dataframe(summary_df, hide_index=True, use_container_width=True)
    else:
        st.info("Import data and generate a plan to see machine loading.")

with plan_tab:
    st.subheader("Generated Production Plan")
    plan_df = rows_to_dataframe(db.plan_rows())
    if plan_df.empty:
        st.info("No production plan has been generated.")
    else:
        preferred = [
            "plan_id", "schedule_id", "customer_name", "shift_name", "machine_name",
            "operation_name", "part_name", "planned_qty", "start_datetime",
            "end_datetime", "process_sequence", "process_type", "due_datetime", "note",
        ]
        columns = [col for col in preferred if col in plan_df.columns]
        st.dataframe(plan_df[columns], hide_index=True, use_container_width=True, height=520)
        st.download_button(
            "Download Production Plan Excel",
            data=export_plan_bytes(),
            file_name=f"Electro_Dip_Production_Plan_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with slips_tab:
    st.subheader("Operator Machine Slip Updates")
    st.info(
        "Enter Actual Qty and Rejected Qty. The accepted Good Qty is Actual minus Rejected. "
        "When the next production plan is generated, accepted quantity at the final in-house "
        "operation is deducted automatically. The revised plan uses the exact remaining "
        "balance, and the final transportation lot may be smaller."
    )
    selected_date = st.date_input(
        "Slip date",
        value=st.session_state.selected_slip_date,
        key="slip_date_widget",
    )
    st.session_state.selected_slip_date = selected_date
    slip_rows = db.todays_slips(selected_date)
    slip_df = rows_to_dataframe(slip_rows)
    if slip_df.empty:
        st.info(f"No in-house operations found for {selected_date:%d-%b-%Y}.")
    else:
        edit_cols = [
            "plan_id", "machine_name", "shift_name", "customer_name",
            "operation_name", "part_name", "planned_qty", "actual_qty",
            "rejected_qty", "status", "operator_name", "supervisor_name",
            "remarks", "start_datetime", "end_datetime", "schedule_id",
            "process_sequence", "priority",
        ]
        editor_df = slip_df[[c for c in edit_cols if c in slip_df.columns]].copy()
        edited_df = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            height=480,
            disabled=[
                "plan_id", "machine_name", "shift_name", "customer_name",
                "operation_name", "part_name", "planned_qty",
                "start_datetime", "end_datetime", "schedule_id",
                "process_sequence", "priority",
            ],
            column_config={
                "actual_qty": st.column_config.NumberColumn("Actual Qty", min_value=0.0, step=1.0),
                "rejected_qty": st.column_config.NumberColumn("Rejected Qty", min_value=0.0, step=1.0),
                "status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Not Started", "Running", "Completed", "Hold", "Rework"],
                ),
                "operator_name": st.column_config.TextColumn("Operator"),
                "supervisor_name": st.column_config.TextColumn("Supervisor"),
                "remarks": st.column_config.TextColumn("Remarks"),
            },
            key=f"slip_editor_{selected_date.isoformat()}",
        )
        c1, c2 = st.columns([1, 2])
        with c1:
            if st.button("Save Operator Production Update", type="primary", use_container_width=True):
                records = edited_df.fillna("").to_dict("records")
                saved = db.save_production_updates(records, selected_date)
                st.success(f"{saved} operation update(s) saved. Subsequent planning will use this progress.")
                st.rerun()
        with c2:
            st.caption(
                "Only accepted output from the last in-house process is deducted from the schedule, "
                "preventing the same quantity from being counted at every intermediate operation."
            )

        pdf_data = slip_pdf_bytes(selected_date)
        st.download_button(
            "Download Approved Slip PDF",
            data=pdf_data,
            file_name=f"Machine_Slips_{selected_date.isoformat()}.pdf",
            mime="application/pdf",
        )

with progress_tab:
    st.subheader("Production Progress Register")
    progress_df = rows_to_dataframe(db.production_progress_rows())
    if progress_df.empty:
        st.info("No operator production updates have been saved.")
    else:
        show_cols = [
            "report_date", "schedule_id", "customer_name", "part_name",
            "operation_name", "machine_name", "planned_qty", "actual_qty",
            "rejected_qty", "good_qty", "status", "operator_name",
            "supervisor_name", "remarks", "updated_at",
        ]
        st.dataframe(
            progress_df[[c for c in show_cols if c in progress_df.columns]],
            hide_index=True,
            use_container_width=True,
            height=540,
        )
        st.download_button(
            "Download Production Progress CSV",
            data=progress_df.to_csv(index=False).encode("utf-8"),
            file_name=f"Production_Progress_{date.today().isoformat()}.csv",
            mime="text/csv",
        )

with data_tab:
    st.subheader("Imported Master Data")
    data_choice = st.selectbox(
        "Select data table",
        [
            "customer_schedules", "stock_demand", "batch_config", "process_bom",
            "machine_recommendations", "machine_downtime", "shifts", "breaks",
            "holidays", "weekly_offs",
        ],
    )
    master_df = pd.read_sql_query(f"SELECT * FROM {data_choice} LIMIT 10000", db.conn)
    st.dataframe(master_df, hide_index=True, use_container_width=True, height=520)

st.divider()
st.caption(
    "Session data is isolated in a temporary database on the server. For permanent multi-user storage, "
    "connect the app to PostgreSQL or another managed database."
)
