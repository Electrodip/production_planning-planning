
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

logo_path = BASE_DIR / "electro_dip_logo.png"

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 0.65rem;
            padding-bottom: 1.5rem;
            max-width: 98%;
        }

        .ed-topbar {
            background: linear-gradient(90deg, #173B63 0%, #245B8E 60%, #2F75B5 100%);
            border-radius: 14px;
            padding: 14px 18px;
            min-height: 84px;
            box-shadow: 0 6px 18px rgba(31, 78, 120, 0.18);
            border: 1px solid rgba(255,255,255,0.15);
        }

        .ed-company {
            color: #FFFFFF;
            font-size: 27px;
            font-weight: 800;
            line-height: 1.05;
            letter-spacing: 0.4px;
            margin-bottom: 4px;
        }

        .ed-system {
            color: #EAF3FB;
            font-size: 14px;
            font-weight: 700;
            margin-bottom: 3px;
        }

        .ed-subline {
            color: #D0E3F3;
            font-size: 11px;
            line-height: 1.4;
        }

        .ed-kpi-card {
            background: #FFFFFF;
            border: 1px solid #D9E4EE;
            border-radius: 13px;
            min-height: 84px;
            padding: 12px 10px;
            box-shadow: 0 4px 12px rgba(31, 78, 120, 0.08);
            text-align: center;
        }

        .ed-kpi-card.blue { border-top: 4px solid #2F75B5; }
        .ed-kpi-card.green { border-top: 4px solid #70AD47; }
        .ed-kpi-card.orange { border-top: 4px solid #ED7D31; }
        .ed-kpi-card.red { border-top: 4px solid #C00000; }

        .ed-kpi-label {
            color: #667788;
            font-size: 10px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.45px;
            margin-bottom: 5px;
        }

        .ed-kpi-value {
            color: #18232D;
            font-size: 25px;
            font-weight: 800;
            line-height: 1;
        }

        .ed-kpi-note {
            color: #8A98A5;
            font-size: 9px;
            margin-top: 6px;
        }

        .ed-section-label {
            margin: 10px 0 7px 2px;
            color: #173B63;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 0.45px;
            text-transform: uppercase;
        }

        div[data-testid="stFileUploader"] {
            border: 1px dashed #AFC4D8;
            border-radius: 10px;
            padding: 3px 6px;
            background: #FBFDFF;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 12px;
            border-color: #D9E4EE !important;
            box-shadow: 0 3px 10px rgba(31, 78, 120, 0.05);
        }

        div[data-baseweb="tab-list"] {
            background: #FFFFFF;
            border: 1px solid #D9E2EA;
            border-radius: 12px;
            padding: 6px 7px 2px 7px;
            gap: 2px;
            box-shadow: 0 3px 10px rgba(31, 78, 120, 0.07);
            overflow-x: auto;
        }

        button[data-baseweb="tab"] {
            border-radius: 8px 8px 0 0;
            padding-left: 12px !important;
            padding-right: 12px !important;
            font-weight: 650 !important;
            color: #30475C !important;
            white-space: nowrap;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            color: #C00000 !important;
            background: #FFF3F1 !important;
        }

        div[data-baseweb="tab-highlight"] {
            background-color: #C00000 !important;
            height: 3px !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid #DDE5EC;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 3px 10px rgba(31, 78, 120, 0.05);
        }

        @media (max-width: 1050px) {
            .ed-company { font-size: 23px; }
            .ed-system { font-size: 12px; }
            .ed-subline { font-size: 10px; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

header_schedule_count = db.conn.execute(
    "SELECT COUNT(*) FROM customer_schedules"
).fetchone()[0]
header_plan_count = db.conn.execute(
    "SELECT COUNT(*) FROM production_plan"
).fetchone()[0]
header_entry_count = db.conn.execute(
    "SELECT COUNT(*) FROM operator_entries"
).fetchone()[0]
header_wip_total = sum(
    float(row["wip_after_process"] or 0)
    for row in db.wip_rows()
)

brand, k1, k2, k3, k4 = st.columns([2.55, 1, 1, 1, 1], gap="small")

with brand:
    logo_col, text_col = st.columns([0.28, 1.72], gap="small")
    with logo_col:
        if logo_path.exists():
            st.image(str(logo_path), width=68)
    with text_col:
        st.markdown(
            """
            <div class="ed-topbar">
                <div class="ed-company">ELECTRO-DIP</div>
                <div class="ed-system">Production Planning & WIP Control System</div>
                <div class="ed-subline">
                    Backward scheduling • Machine-wise slips • Persistent entries •
                    Process-wise WIP reporting
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

with k1:
    st.markdown(
        f"""
        <div class="ed-kpi-card blue">
            <div class="ed-kpi-label">Schedules</div>
            <div class="ed-kpi-value">{header_schedule_count:,}</div>
            <div class="ed-kpi-note">Customer lines</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="ed-kpi-card green">
            <div class="ed-kpi-label">Plan Rows</div>
            <div class="ed-kpi-value">{header_plan_count:,}</div>
            <div class="ed-kpi-note">Operations</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    st.markdown(
        f"""
        <div class="ed-kpi-card orange">
            <div class="ed-kpi-label">Saved Entries</div>
            <div class="ed-kpi-value">{header_entry_count:,}</div>
            <div class="ed-kpi-note">Operator updates</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    st.markdown(
        f"""
        <div class="ed-kpi-card red">
            <div class="ed-kpi-label">Physical WIP</div>
            <div class="ed-kpi-value">{header_wip_total:,.0f}</div>
            <div class="ed-kpi-note">Current WIP qty</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="ed-section-label">Data Import</div>', unsafe_allow_html=True)

with st.container(border=True):
    upload_col, template_col, import_col, refresh_col = st.columns(
        [2.4, 1, 1, 0.8],
        gap="small",
    )

    with upload_col:
        uploaded_excel = st.file_uploader(
            "Upload Production Planning Excel",
            type=["xlsx", "xlsm"],
            key="master_excel_upload",
            help="Use the Electro-Dip import template.",
            label_visibility="collapsed",
        )

    with template_col:
        if TEMPLATE_PATH.exists():
            st.download_button(
                "Download Template",
                data=TEMPLATE_PATH.read_bytes(),
                file_name=TEMPLATE_PATH.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with import_col:
        import_clicked = st.button(
            "Import Excel",
            type="primary",
            use_container_width=True,
        )

    with refresh_col:
        if st.button("Refresh", use_container_width=True):
            st.rerun()

    if uploaded_excel is not None:
        st.caption(
            f"Selected: {uploaded_excel.name} "
            f"({uploaded_excel.size / 1024:.1f} KB)"
        )

    if import_clicked:
        if uploaded_excel is None:
            st.error("Select an Excel file before importing.")
        else:
            try:
                import_report = db.import_workbook(uploaded_excel)
                st.session_state["latest_import_report"] = import_report
                st.success(
                    "Excel imported successfully. "
                    f"Previous entries preserved: "
                    f"{import_report['previous_entries_preserved']}"
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Excel import failed: {exc}")

    if "latest_import_report" in st.session_state:
        latest_report = st.session_state["latest_import_report"]
        with st.expander("Latest Import Report"):
            counts = latest_report.get("counts", {})
            if counts:
                counts_df = pd.DataFrame(
                    [{"Data Type": k, "Imported Rows": v} for k, v in counts.items()]
                )
                st.dataframe(counts_df, hide_index=True, use_container_width=True)

            warnings = latest_report.get("warnings", [])
            if warnings:
                for warning in warnings[:100]:
                    st.write("•", warning)
            else:
                st.success("No import warnings.")

st.markdown('<div class="ed-section-label">Application Modules</div>', unsafe_allow_html=True)

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
    st.info(
        "V14 schedule quantity remains unchanged. Transportation batch only "
        "splits the exact process quantity into lot rows. Example: "
        "240 with batch 100 becomes 100 + 100 + 40."
    )
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
        preferred_plan_columns = [
            "plan_id", "schedule_id", "customer_name", "part_name",
            "process_sequence", "operation_name", "machine_name",
            "shift_name", "schedule_plan_qty_v14",
            "net_process_plan_qty", "transportation_batch_qty",
            "lot_no", "lot_planned_qty",
            "production_start_datetime", "production_end_datetime",
            "due_datetime", "priority", "note",
        ]
        visible_plan_columns = [
            column for column in preferred_plan_columns
            if column in plan_df.columns
        ]
        st.dataframe(
            plan_df[visible_plan_columns],
            hide_index=True,
            use_container_width=True,
            height=570,
        )
        st.caption(
            "Schedule Plan Qty (V14) is the approved schedule calculation. "
            "Net Process Plan Qty is after process-wise WIP. Lot Planned Qty "
            "is the transportation-batch split; all lot totals equal the exact "
            "net process quantity."
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
                "plan_id", "machine_name", "customer_name", "part_name",
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

    plan_rows = db.plan_id_rows()
    if not plan_rows:
        st.warning("Generate the production plan first.")
    else:
        plan_ids = [row["plan_id"] for row in plan_rows]

        top1, top2 = st.columns([2, 1])
        selected_plan_id = top1.selectbox(
            "Plan ID",
            plan_ids,
            key="operator_plan_id",
        )
        plan_detail = db.plan_id_detail(selected_plan_id)
        gate_status = db.sequence_gate_status(selected_plan_id)

        with top2:
            st.caption("Selected plan")
            st.write(
                f"{plan_detail['part_name']} | "
                f"{plan_detail['operation_name']} | "
                f"{plan_detail.get('machine_name') or '-'} | "
                f"{plan_detail.get('shift_name') or '-'} | "
                f"Qty {plan_detail['planned_qty']:g}"
            )

        part = str(plan_detail["part_name"])
        sequence = int(plan_detail["process_sequence"])
        operation = str(plan_detail["operation_name"])
        customer_name = str(plan_detail.get("customer_name") or "")
        schedule_id = str(plan_detail.get("schedule_id") or "")
        planned_default = float(plan_detail.get("planned_qty") or 0)

        production_start = plan_detail.get("production_start_datetime")
        production_date = (
            pd.to_datetime(production_start).date()
            if production_start else date.today()
        )

        machine = str(plan_detail.get("machine_name") or "")
        shift = str(plan_detail.get("shift_name") or "")
        production_end = plan_detail.get("production_end_datetime")

        start_time_text = (
            pd.to_datetime(production_start).strftime("%H:%M")
            if production_start else ""
        )
        end_time_text = (
            pd.to_datetime(production_end).strftime("%H:%M")
            if production_end else ""
        )

        c1, c2 = st.columns(2)
        c1.text_input("Part Number", value=part, disabled=True)
        c2.text_input(
            "Process / Sequence",
            value=f"{sequence} - {operation}",
            disabled=True,
        )

        c3, c4, c5 = st.columns(3)
        c3.text_input(
            "Production Date",
            value=production_date.strftime("%Y/%m/%d"),
            disabled=True,
        )
        c4.text_input("Machine", value=machine, disabled=True)
        c5.text_input("Shift", value=shift, disabled=True)

        c6, c7, c8 = st.columns(3)
        c6.text_input(
            "Production Start Time",
            value=start_time_text,
            disabled=True,
        )
        c7.text_input(
            "Production End Time",
            value=end_time_text,
            disabled=True,
        )
        c8.text_input(
            "Planned Qty",
            value=f"{planned_default:g}",
            disabled=True,
        )

        st.markdown("#### Operation Sequence Control")
        g1, g2, g3, g4 = st.columns(4)
        g1.metric("Lot Planned Qty", f"{gate_status['lot_planned_qty']:g}")
        g2.metric(
            "Previous Operation Good Qty",
            f"{gate_status['previous_good_qty']:g}",
        )
        g3.metric(
            "Already Processed Here",
            f"{gate_status['already_processed_here']:g}",
        )
        g4.metric(
            "Maximum Entry Allowed",
            f"{gate_status['maximum_entry_allowed']:g}",
        )

        if gate_status["gate_open"]:
            st.success(gate_status["reason"])
        else:
            st.error(gate_status["reason"])

        c9, c10 = st.columns(2)
        actual = c9.number_input(
            "Actual Qty",
            min_value=0.0,
            max_value=float(gate_status["maximum_entry_allowed"]),
            step=1.0,
            disabled=not gate_status["gate_open"],
            key=f"actual_{selected_plan_id}",
        )
        rejected = c10.number_input(
            "Rejected Qty",
            min_value=0.0,
            max_value=float(actual),
            step=1.0,
            disabled=not gate_status["gate_open"],
            key=f"rejected_{selected_plan_id}",
        )

        entry_date = production_date
        planned = planned_default

        operator_names = db.personnel_names("OPERATOR")
        supervisor_names = db.personnel_names("SUPERVISOR")

        p1, p2 = st.columns(2)
        operator = p1.selectbox(
            "Operator Name",
            operator_names if operator_names else [""],
            key=f"operator_{selected_plan_id}",
        )
        supervisor = p2.selectbox(
            "Supervisor Name",
            supervisor_names if supervisor_names else [""],
            key=f"supervisor_{selected_plan_id}",
        )

        with st.expander("Add Operator / Supervisor Name"):
            a1, a2, a3 = st.columns([2, 1, 1])
            new_person_name = a1.text_input("Name")
            new_person_role = a2.selectbox(
                "Role", ["OPERATOR", "SUPERVISOR"]
            )
            if a3.button("Add Name"):
                try:
                    db.add_personnel(new_person_name, new_person_role)
                    st.success("Name added. Refresh this tab to see it in the dropdown.")
                except Exception as exc:
                    st.error(str(exc))

        status = st.selectbox(
            "Status",
            ["Running", "Completed", "Hold", "Rework"],
            key=f"status_{selected_plan_id}",
        )
        remarks = st.text_area(
            "Remarks",
            key=f"remarks_{selected_plan_id}",
        )

        if st.button(
            "Save Operator Entry",
            type="primary",
            disabled=not gate_status["gate_open"],
        ):
            if rejected > actual:
                st.error("Rejected Qty cannot exceed Actual Qty.")
            elif not operator:
                st.error("Select an Operator Name.")
            elif not supervisor:
                st.error("Select a Supervisor Name.")
            else:
                try:
                    good = db.add_operator_entry(
                        entry_date=entry_date,
                        part_name=part,
                        process_sequence=sequence,
                        operation_name=operation,
                        actual_qty=actual,
                        rejected_qty=rejected,
                        machine_name=machine,
                        shift_name=shift,
                        plan_id=selected_plan_id,
                        schedule_id=schedule_id,
                        customer_name=customer_name,
                        planned_qty=planned,
                        status=status,
                        operator_name=operator,
                        supervisor_name=supervisor,
                        remarks=remarks,
                    )
                    st.success(
                        f"Entry saved for Plan ID {selected_plan_id}. "
                        f"Good Qty = {good:g}."
                    )
                except Exception as exc:
                    st.error(str(exc))

        today_entries = pd.DataFrame([
            dict(r) for r in db.operator_entries()
            if str(r["entry_date"]) == entry_date.isoformat()
        ])
        st.markdown(f"#### Entries for Production Date ({entry_date:%d-%b-%Y})")
        if today_entries.empty:
            st.info("No entries for the selected date.")
        else:
            preferred = [
                "plan_id", "part_name", "process_sequence",
                "operation_name", "machine_name", "shift_name",
                "planned_qty", "actual_qty", "rejected_qty",
                "operator_name", "supervisor_name", "status",
                "created_at",
            ]
            visible = [c for c in preferred if c in today_entries.columns]
            st.dataframe(
                today_entries[visible],
                hide_index=True,
                use_container_width=True,
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
