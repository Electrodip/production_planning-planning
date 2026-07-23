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


## Version 10 complete modules

- Production Plan screen with Generate / Regenerate button
- Downloadable Production Plan Excel
- Machine-wise Operator Slips tab
- Approved ELECTRO-DIP operator-slip PDF
- One slip page per machine and selected due date
- Multiple planned operations on one slip
- Actual Qty and Rejected Qty columns on every slip row
- Existing persistent Operator Entry and all WIP reports retained


## Version 11 timed production plan

Production Plan and Operator Slips now include:
- Production Start Date and Time
- Production End Date and Time
- Shift Name
- Backward scheduling from customer due date
- Cycle time and setup time
- Shift calendar and breaks
- Holidays and weekly offs
- Machine downtime
- Recommended-machine capacity selection
- Outsource lead-time scheduling


## Version 12 dropdown and Plan ID update

Daily Operator Entry now includes:
- Plan ID dropdown
- Machine dropdown
- Shift dropdown
- Operator dropdown
- Supervisor dropdown
- Add-name personnel master
- Automatic part/process/qty/date selection from Plan ID
- Plan ID saved in operator history
- Plan ID printed on operator-slip PDF


## Version 13 — Plan ID in every slip row

The operator-slip table now prints Plan ID as the first column.
Every subsequent operation row displays its own Plan ID.


## Version 14 — Automatic Plan Fields

Machine, Shift, Production Date, Production Start Time,
Production End Time and Planned Qty are automatically populated
from the selected Plan ID and are read-only.
