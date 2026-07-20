# Electro-Dip Production Planner Desktop App

## Main features

- Excel import using the supplied template
- Customer schedule and minimum-stock demand
- Production batch and transportation batch planning
- Process BOM and process sequence
- Up to 15 recommended machines per operation in the template
- Normalized database design, allowing more machines later
- Backward production scheduling
- Shared-machine conflict checking
- Shifts, lunch/tea breaks, holidays and weekly offs
- Machine downtime
- Outsource lead times
- Production plan viewer
- Machine-wise slips by selected date
- Excel export
- Clear-plan confirmation and row-count message
- Local SQLite database

## Run from Python

1. Install Python 3.11 or later.
2. Open the app folder.
3. Run `pip install -r requirements.txt`.
4. Double-click `Run_Planner.bat`.

## Build Windows EXE

Double-click `Build_EXE.bat`.

The executable will be created in:

`dist/ElectroDipProductionPlanner.exe`

## Workflow

1. Open the app.
2. Click **Import Excel** and choose `Electro_Dip_Desktop_App_Import_Template.xlsx`.
3. Click **Validate**.
4. Click **Generate Plan**.
5. Review **Production Plan** and **Machine Slips**.
6. Click **Export Plan**.

## Important planning rule

The app chooses the recommended machine providing the latest feasible backward slot.
When start times are equal, the lower recommendation order has priority.


## Small Decorated Machine Slip Printing

The app now includes:

- **Create Slip PDF**: creates an A6 landscape decorated PDF.
- **Print Small Slips**: sends the slip PDF to the Windows default printer.
- One machine operation per slip/page.
- Suitable paper sizes:
  - A6 landscape
  - 4 x 6 inch card
- Printed fields:
  - Machine name
  - Slip number
  - Date and shift
  - Part and operation
  - Planned quantity
  - Start and end time
  - Schedule ID and priority
  - Supervisor instruction
  - Actual quantity
  - Status
  - Operator signature

Select the required date in the **Machine Slips** tab before printing.


## Small Coloured Machine Slips

The Machine Slips tab now has:

- **Create Coloured PDF** - creates and opens an A6 landscape PDF.
- **Print Small Slips** - sends the PDF to the default Windows printer.
- One operation is printed per small slip/page.
- Recommended paper size: **A6 landscape (148 x 105 mm)**.
- For thermal printers, choose a compatible A6/4x6-inch page size in printer settings.

Each slip includes machine, part, operation, planned quantity, shift,
start/end time, schedule ID, process sequence, supervisor instruction,
actual quantity, rejection quantity, status boxes, and signatures.


## Version 3 changes

- Company heading changed to **ELECTRO-DIP** only.
- Customer Name added to the Customer_Schedules import sheet.
- Customer Name is stored in each production-plan row.
- Customer is shown in each operation row on the printed machine slip.
- Customer is not shown in the slip header.
- One slip groups all operations by Machine + Shift + Selected Date.
- Actual Qty and Rejected Qty are available in every operation row.
- Up to 15 recommended machines are supported in the import template.


## Approved printing format

The application prints only the approved ELECTRO-DIP Machine Daily Production Slip:

- A4 landscape
- one page per selected Date + Machine + Shift
- ELECTRO-DIP logo and heading
- Customer shown only in each operation row
- multiple operations on the same page
- Plan Qty, Actual Qty and Rejected Qty columns
- Priority, Status and Remarks columns
- totals, utilization, supervisor/quality remarks
- Completed / Hold / Running / Rework status
- Operator / Quality / Supervisor signatures


## Version 5 robust importer

- Blank formatted Excel rows are ignored automatically.
- Machine_Downtime rows containing only a machine name are treated as
  placeholders and skipped without an error.
- Incomplete records are skipped and reported instead of crashing the import.
- Import results display record counts and a short warning summary.
- Validation and planning dialogs show a maximum of 30 detailed issues,
  with the remaining issue count summarized.
