# ELECTRO-DIP Online Production Planner

## Features

- Browser-based Excel upload
- Robust import: blank rows and downtime placeholders are skipped
- Import counts and warning summary
- Validation before planning
- Backward production scheduling
- 15 recommended machines per operation
- Machine conflict, shifts, breaks, holidays, weekly offs and downtime
- Production plan dashboard and Excel download
- Approved grouped Machine Daily Production Slip PDF
- Customer shown in each operation row
- Optional password protection

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

On Windows, double-click `run_online_app.bat`.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload all files from this folder.
3. In Streamlit Community Cloud, create an app and select `streamlit_app.py` as the entrypoint.
4. Optional: add a private app password in **App settings > Secrets**:

```toml
APP_PASSWORD = "your-private-password"
```

5. Deploy.

## Important storage note

The included version stores each browser session in an isolated temporary SQLite database. This is suitable for testing and private single-session planning. For permanent multi-user production use, replace temporary SQLite storage with PostgreSQL or another managed database.

## Printing

Go to **Machine Slips**, select the date, and download the approved PDF. The PDF groups all operations by Date + Machine + Shift and creates one A4 landscape page per group.

## V2 SQLite thread fix

This build does not cache a live SQLite connection with `st.cache_resource`.
Each Streamlit rerun opens a new connection to the same session-specific
SQLite database file. This prevents `sqlite3.ProgrammingError` when Streamlit
executes a rerun on another thread.
