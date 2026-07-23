
import os
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from planner_core import Database


APP_TITLE = "Electro-Dip Production Planner with WIP"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "electro_dip_persistent.db"
TEMPLATE_PATH = BASE_DIR / "Electro_Dip_WIP_Import_Template_V7.xlsx"

st.set_page_config(page_title=APP_TITLE, page_icon="🏭", layout="wide")


def get_db():
    return Database(DB_PATH)



db = get_db()


def make_operator_slip_pdf_bytes(selected_date=None, selected_machine=None):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        path = Path(tmp.name)
    try:
        db.create_operator_slips_pdf(
            path,
            due_date=selected_date,
            machine_name=selected_machine,
        )
        return path.read_bytes()
    finally:
        path.unlink(missing_ok=True)


def make_production_plan_excel_bytes():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = Path(tmp.name)
    try:
        db.export_production_plan_excel(path)
        return path.read_bytes()
    finally:
        path.unlink(missing_ok=True)

st.title("ELECTRO-DIP")
st.subheader("Online Production Planning System with Persistent WIP")
st.caption(
    "Operator entries and WIP remain saved through imports and plan regeneration. "
    "They are deleted only from the Clear Data tab."
)

with st.sidebar:
    st.header("Import Master Data")
    uploaded = st.file_uploader("Upload Excel template", type=["xlsx", "xlsm"])
    if TEMPLATE_PATH.exists():
        st.download_button(
            "Download WIP Import Template",
            TEMPLATE_PATH.read_bytes(),
            file_name=TEMPLATE_PATH.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    if st.button("Import Excel", type="primary", use_container_width=True):
        if uploaded is None:
            st.error("Select an Excel file first.")
        else:
            report = db.import_workbook(uploaded)
            st.session_state["import_report"] = report
            st.success(
                f"Imported successfully. Previous entries preserved: "
                f"{report['previous_entries_preserved']}"
            )

    if st.button("Generate / Regenerate Plan", use_container_width=True):
        count = db.generate_plan()
        st.success(f"Production plan generated: {count} rows")

    if st.button("Clear Current Plan Only", use_container_width=True):
        count = db.clear_plan()
        st.info(f"{count} plan rows cleared. Operator entries and WIP were preserved.")

if "import_report" in st.session_state:
    with st.expander("Latest Import Report"):
        report = st.session_state["import_report"]
        st.json(report["counts"])
        if report["warnings"]:
            for warning in report["warnings"][:50]:
                st.write("•", warning)

tabs = st.tabs([
    "Dashboard", "Production Plan", "Operator Slips", "Operator Entry",
    "Process WIP Report", "WIP Ageing", "WIP Summaries",
    "Schedule Calculation", "Process WIP Allocation",
    "Previous Entries", "Download Excel", "Clear Data"
])

with tabs[0]:
    entries = db.conn.execute("SELECT COUNT(*) FROM operator_entries").fetchone()[0]
    wip_total = sum(r["wip_after_process"] for r in db.wip_rows())
    dispatched = db.conn.execute(
        """SELECT COALESCE(SUM(e.good_qty),0)
           FROM operator_entries e
           WHERE e.process_sequence = (
               SELECT MAX(p.process_sequence)
               FROM process_bom p WHERE p.part_name=e.part_name
           )"""
    ).fetchone()[0]
    plan_rows = db.conn.execute("SELECT COUNT(*) FROM production_plan").fetchone()[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Saved Entries", entries)
    c2.metric("Physical WIP Qty", round(wip_total, 2))
    c3.metric("Dispatched Good Qty", round(float(dispatched or 0), 2))
    c4.metric("Plan Rows", plan_rows)
    st.info(
        "Highest Process BOM sequence is treated as Dispatch. "
        "Dispatch quantity is allocated FIFO to the earliest due schedule. "
        "Minimum stock is generated once, after customer schedules."
    )

with tabs[1]:
    st.subheader("Production Plan")
    p1, p2 = st.columns([1, 1])
    with p1:
        if st.button("Generate / Regenerate Production Plan", type="primary"):
            count = db.generate_plan()
            st.success(f"Production plan generated: {count} rows")
    with p2:
        plan_rows_for_download = db.production_plan_report_rows()
        if plan_rows_for_download:
            st.download_button(
                "Download Production Plan Excel",
                data=make_production_plan_excel_bytes(),
                file_name="Electro_Dip_Production_Plan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    plan_df = pd.DataFrame(db.production_plan_report_rows())
    if plan_df.empty:
        st.info(
            "No production plan is available. Import the Excel template and "
            "click Generate / Regenerate Production Plan."
        )
    else:
        st.dataframe(
            plan_df,
            hide_index=True,
            use_container_width=True,
            height=570,
        )
        st.caption(
            "The plan is WIP-adjusted, transportation-batch split, and "
            "backward scheduled using shifts, breaks, holidays, weekly offs, "
            "machine downtime, cycle time and setup time."
        )

with tabs[2]:
    st.subheader("Machine-wise Operator Slips")
    slip_dates = db.operator_slip_dates()
    if not slip_dates:
        st.info("Generate the production plan before creating operator slips.")
    else:
        s1, s2 = st.columns(2)
        selected_slip_date = s1.selectbox(
            "Production Date",
            slip_dates,
            format_func=lambda d: d.strftime("%d-%b-%Y"),
        )
        machines = ["All Machines"] + db.operator_slip_machines(selected_slip_date)
        selected_machine_label = s2.selectbox("Machine", machines)
        selected_machine = (
            None if selected_machine_label == "All Machines"
            else selected_machine_label
        )

        slip_rows = db.operator_slip_rows(
            due_date=selected_slip_date,
            machine_name=selected_machine,
        )
        slip_df = pd.DataFrame(slip_rows)
        if not slip_df.empty:
            preferred = [
                "machine_name", "customer_name", "part_name",
                "process_sequence", "operation_name", "planned_qty",
                "production_start_datetime", "production_end_datetime",
                "shift_name", "lot_no", "schedule_id",
                "due_datetime", "priority"
            ]
            visible = [c for c in preferred if c in slip_df.columns]
            st.dataframe(
                slip_df[visible],
                hide_index=True,
                use_container_width=True,
                height=440,
            )
            st.download_button(
                "Download / Print Operator Slip PDF",
                data=make_operator_slip_pdf_bytes(
                    selected_slip_date,
                    selected_machine,
                ),
                file_name=(
                    f"Operator_Slips_{selected_slip_date.isoformat()}.pdf"
                ),
                mime="application/pdf",
                type="primary",
            )
            st.caption(
                "One page is generated per machine. Multiple operations are "
                "printed on the same approved ELECTRO-DIP slip."
            )

with tabs[3]:
    st.subheader("Daily Operator Entry")
    bom_df = pd.DataFrame([dict(r) for r in db.bom_rows()])
    if bom_df.empty:
        st.warning("Import Process BOM first.")
    else:
        parts = sorted(bom_df["part_name"].astype(str).unique())
        part = st.selectbox("Part Number", parts)
        part_ops = bom_df[bom_df["part_name"].astype(str) == str(part)]
        sequence = st.selectbox(
            "Process / Sequence",
            part_ops["process_sequence"].tolist(),
            format_func=lambda seq: (
                f"{int(seq)} - "
                f"{part_ops[part_ops['process_sequence']==seq]['operation_name'].iloc[0]}"
            ),
        )
        operation = part_ops[
            part_ops["process_sequence"] == sequence
        ]["operation_name"].iloc[0]

        c1, c2, c3 = st.columns(3)
        entry_date = c1.date_input("Entry Date", value=date.today())
        machine = c2.text_input("Machine")
        shift = c3.text_input("Shift")

        c4, c5, c6 = st.columns(3)
        actual = c4.number_input("Actual Qty", min_value=0.0, step=1.0)
        rejected = c5.number_input("Rejected Qty", min_value=0.0, step=1.0)
        planned = c6.number_input("Planned Qty", min_value=0.0, step=1.0)

        c7, c8 = st.columns(2)
        operator = c7.text_input("Operator Name")
        supervisor = c8.text_input("Supervisor Name")
        status = st.selectbox(
            "Status", ["", "Running", "Completed", "Hold", "Rework"]
        )
        remarks = st.text_area("Remarks")

        if st.button("Save Operator Entry", type="primary"):
            if rejected > actual:
                st.error("Rejected Qty cannot exceed Actual Qty.")
            else:
                good = db.add_operator_entry(
                    entry_date=entry_date,
                    part_name=part,
                    process_sequence=int(sequence),
                    operation_name=operation,
                    actual_qty=actual,
                    rejected_qty=rejected,
                    machine_name=machine,
                    shift_name=shift,
                    planned_qty=planned,
                    status=status,
                    operator_name=operator,
                    supervisor_name=supervisor,
                    remarks=remarks,
                )
                st.success(
                    f"Entry saved. Good Qty = {good:g}. "
                    "This entry will remain until Clear Previous Entries is used."
                )


with tabs[4]:
    st.subheader("Process-wise WIP Report")
    report_df = pd.DataFrame(db.process_wip_report_rows())
    if report_df.empty:
        st.info("No WIP data.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        customers = ["All"] + sorted([x for x in report_df["customer_name"].dropna().astype(str).unique() if x])
        parts = ["All"] + sorted(report_df["part_name"].astype(str).unique())
        processes = ["All"] + sorted(report_df["operation_name"].astype(str).unique())
        machines = ["All"] + sorted([x for x in report_df["machine_name"].dropna().astype(str).unique() if x])
        customer = f1.selectbox("Customer", customers)
        part = f2.selectbox("Part", parts)
        process = f3.selectbox("Process", processes)
        machine = f4.selectbox("Machine", machines)

        filtered = report_df.copy()
        if customer != "All":
            filtered = filtered[filtered["customer_name"].astype(str) == customer]
        if part != "All":
            filtered = filtered[filtered["part_name"].astype(str) == part]
        if process != "All":
            filtered = filtered[filtered["operation_name"].astype(str) == process]
        if machine != "All":
            filtered = filtered[filtered["machine_name"].astype(str) == machine]

        st.dataframe(filtered, hide_index=True, use_container_width=True, height=540)
        st.caption("WIP Qty = normalized cumulative good at current process minus normalized cumulative good at next process.")

with tabs[5]:
    st.subheader("WIP Ageing")
    ageing_df = pd.DataFrame(db.process_wip_report_rows())
    ageing_df = ageing_df[ageing_df["wip_qty"] > 0] if not ageing_df.empty else ageing_df
    if ageing_df.empty:
        st.info("No open WIP.")
    else:
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Total WIP", round(float(ageing_df["wip_qty"].sum()), 2))
        a2.metric("Fresh", round(float(ageing_df.loc[ageing_df["age_status"]=="Fresh", "wip_qty"].sum()), 2))
        a3.metric("Monitor", round(float(ageing_df.loc[ageing_df["age_status"]=="Monitor", "wip_qty"].sum()), 2))
        a4.metric("Old WIP", round(float(ageing_df.loc[ageing_df["age_status"]=="Old WIP", "wip_qty"].sum()), 2))
        st.dataframe(
            ageing_df.sort_values(["wip_age_days", "wip_qty"], ascending=[False, False]),
            hide_index=True, use_container_width=True, height=520
        )

with tabs[6]:
    st.subheader("Machine / Part / Customer WIP")
    s1, s2, s3 = st.columns(3)
    with s1:
        st.markdown("#### Machine-wise")
        st.dataframe(pd.DataFrame(db.machine_wip_rows()), hide_index=True, use_container_width=True)
    with s2:
        st.markdown("#### Part-wise")
        st.dataframe(pd.DataFrame(db.part_wip_rows()), hide_index=True, use_container_width=True)
    with s3:
        st.markdown("#### Customer-wise")
        st.dataframe(pd.DataFrame(db.customer_wip_rows()), hide_index=True, use_container_width=True)

with tabs[7]:
    st.subheader("Schedule-Line Plan Quantity Calculation")
    schedule_calc_df = pd.DataFrame(db.schedule_line_calculation_rows())
    if schedule_calc_df.empty:
        st.info("No schedules imported.")
    else:
        st.dataframe(schedule_calc_df, hide_index=True, use_container_width=True, height=520)
        st.caption("First line: Demand + Minimum Stock - Opening Available. Further lines: Demand + Minimum Stock - Previous Schedule Carry-Forward WIP.")

with tabs[8]:
    st.subheader("Process-BOM-Wise WIP Allocation")
    process_wip_df = pd.DataFrame(db.process_schedule_wip_rows())
    if process_wip_df.empty:
        st.info("No process WIP allocation available.")
    else:
        st.dataframe(process_wip_df, hide_index=True, use_container_width=True, height=520)

with tabs[9]:
    st.subheader("Saved Previous Operator Entries")
    entries_df = pd.DataFrame([dict(r) for r in db.operator_entries()])
    if entries_df.empty:
        st.info("No saved entries.")
    else:
        st.dataframe(entries_df, hide_index=True, use_container_width=True, height=560)

with tabs[10]:
    st.subheader("Download Complete WIP Workbook")
    if st.button("Prepare Downloadable Excel"):
        export_path = BASE_DIR / "Electro_Dip_WIP_Report.xlsx"
        db.export_wip_excel(export_path)
        st.session_state["wip_export"] = export_path.read_bytes()
    if "wip_export" in st.session_state:
        st.download_button(
            "Download WIP Excel",
            st.session_state["wip_export"],
            file_name="Electro_Dip_WIP_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    st.write(
        "Workbook sheets include Process WIP Report, WIP Ageing data, Machine WIP, Part WIP, Customer WIP, Dispatch History, Schedule Calculation, Production Plan and Operator Entries."
    )

with tabs[11]:
    st.subheader("Clear Previous Entries")
    st.error(
        "This is the only action that deletes saved operator entries and WIP history."
    )
    confirm = st.text_input('Type exactly: CLEAR ALL ENTRIES')
    if st.button("Clear Previous Entries and Plan", type="primary"):
        if confirm != "CLEAR ALL ENTRIES":
            st.warning("Confirmation text does not match.")
        else:
            count = db.clear_previous_entries()
            st.session_state.pop("wip_export", None)
            st.success(f"{count} previous entries cleared. Production plan also cleared.")
