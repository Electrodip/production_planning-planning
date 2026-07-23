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


## Version 15 — Attractive Header Grid

The app now includes:
- Branded ELECTRO-DIP header with logo
- Production Planning & WIP Control title
- Header summary cards for schedules, plan rows, saved entries and WIP
- Styled navigation grid with highlighted active tab
- Responsive layout for desktop, tablet and mobile
- Improved metric and report-table styling


## Version 16 — Transportation Lot Planning

The V14 schedule quantity logic remains unchanged.

Transportation batch is used only to split the exact process quantity.

Example:
- Schedule Plan Qty: 240
- Transportation Batch: 100
- Production Lots: 100 + 100 + 40
- Total remains 240

Production Plan includes:
- Schedule Plan Qty V14
- Net Process Plan Qty
- Transportation Batch Qty
- Lot No.
- Lot Planned Qty


## Version 17 — Operation Sequence Gate

The app now prevents process jumping and over-processing.

Rules:
- First process cannot exceed its transportation-lot planned quantity.
- Every next process is limited by good quantity released by the immediate
  previous process for the same Schedule ID and Lot No.
- Partial entries are allowed.
- Quantity already processed at the current operation is deducted.
- Validation is enforced both in the screen and in the database save method.
- Dispatch is therefore limited by the good output of the previous operation.


## Version 21 — V14 Interface and No WIP in Schedule Calculation

Interface:
- Restored simple V14-style Streamlit layout
- Clear title, import panel and normal module tabs
- Removed custom header grid

Schedule calculation:
- Schedule Plan Qty = Demand + Minimum Stock - Allocated Current FG Stock
- Opening WIP is not deducted
- Previous-schedule WIP is not deducted
- Process WIP is considered only after schedule quantity calculation
- Transportation batch splits the exact process quantity
- All later features remain included


## Version 22 — Edit and Delete Previous Entries

Previous Entries tab now includes:
- Entry date, part and Plan ID filters
- Entry ID selection
- Edit Actual Qty
- Edit Rejected Qty
- Edit Operator
- Edit Supervisor
- Edit Status
- Edit Remarks
- Save Changes
- Delete Entry with confirmation

Plan-controlled fields cannot be edited:
- Plan ID
- Part
- Operation
- Machine
- Shift
- Planned Qty
- Production date

WIP, ageing and operation-sequence gates recalculate automatically after
editing or deleting an operator entry.


## Version 23 — Import Diagnostics Before Import

New workflow:
1. Upload Excel
2. Validate Template
3. Review worksheet health and detailed diagnostics
4. Import only after validation passes

Checks include:
- Required sheets
- Required column headings
- Duplicate Schedule IDs
- Duplicate stock/batch records
- Mandatory blank cells
- Invalid quantities, dates and times
- Missing Process BOM, stock or batch master
- Invalid process types
- Missing cycle time or outsource lead time
- Missing machine recommendations
- Incomplete machine downtime
- Invalid shift data

The diagnostic report shows severity, sheet, row, column, reason and value.
Critical errors block import. Warnings do not block import.
