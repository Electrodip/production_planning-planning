# Electro-Dip Online Production Planner V7 with WIP

## Core rules

- Highest Process BOM sequence is Dispatch / completion.
- Dispatch good quantity = Actual Qty - Rejected Qty.
- Dispatch is allocated FIFO to the earliest due customer schedule.
- Earlier process entries remain WIP and do not reduce customer demand.
- WIP after a process = cumulative good at that process minus cumulative good at the next process.
- Current FG stock is allocated once per part, FIFO to customer schedules.
- Minimum stock is added once per part after all customer schedules.
- Transportation batch splits every process plan; the final lot can be smaller.
- Operator entries survive Excel re-import and plan regeneration.
- Only Clear Data -> Clear Previous Entries deletes saved entries.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Cloud

Upload all files to GitHub and deploy `streamlit_app.py`.

## Persistence note

The SQLite file is stored beside the application as `electro_dip_persistent.db`.
It survives normal reruns and imports. On Streamlit Community Cloud, a server
restart or redeployment may reset local disk. For guaranteed permanent
multi-user storage, connect the same database layer to PostgreSQL/Supabase.


## Version 8 schedule-line WIP logic

First schedule line: Plan Qty = Demand + Minimum Stock Level - Opening Available Stock/WIP.

Further schedule lines: Plan Qty = Demand + Minimum Stock Level - Previous Schedule Carry-Forward WIP.

Download workbook includes Schedule_Line_Calculation, Process_WIP_Allocation, WIP_Summary, Production_Plan, Operator_Entries and Logic.


## Version 9 reports
- Process-wise WIP with filters
- WIP ageing: Fresh, Monitor, Old WIP
- Machine-wise WIP
- Part-wise WIP
- Customer-wise WIP
- Dispatch history
- All reports included in downloadable Excel
