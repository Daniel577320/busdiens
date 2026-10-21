# Fleet Management System

A client-facing web application for managing a bus fleet's daily trip records, fuel tanks, service history, and performance analytics. Built with **FastAPI** (Python) on the backend, **SQLite** for storage, and **Jinja2 + vanilla JavaScript + Chart.js** for the frontend — no build step required.

## What it does

- **Home / Entry form** — Log daily trip records per bus and driver, including trip type, kilometres travelled, and fuel usage. Multi-day trip types (e.g. tour-with-overnight, tour-return) automatically fan out into the correct number of calendar-day entries.
- **Dagboek (diary/schedule)** — An inline-editable grid for scheduling and tracking trips and quotations (Kwotasie), with history, filtering (by acceptance status and trip type), and an Excel export.
- **Admin panel** — Tabbed interface to manage:
  - **Buses** — fleet roster, seater capacity, target km/L, and tank assignment
  - **Drivers** — driver roster
  - **Trip Records** — edit/delete historical entries
  - **Service** — service and maintenance history logging per bus
  - **Analytics** — fleet overview, bus performance, driver performance, and fuel tank status, visualized with Chart.js
- **Fuel tanks** — Track refuelling against a shared tank capacity, with admin views into tank status.
- **Good/Bad performance flag** — Each bus's actual km/L is compared against its configured target to flag under/over-performing trips.

## Tech stack

- **Backend:** FastAPI (Python), raw `sqlite3` queries (no ORM)
- **Database:** SQLite (`Database.db`, created via `init_db.py`)
- **Frontend:** Jinja2 templates, vanilla JavaScript, Chart.js
- **Reporting:** Excel export via `openpyxl`

## Project structure

```
.
├── main.py              # FastAPI app: all routes (entry, dagboek, admin, analytics, service, tanks)
├── init_db.py            # Creates/initializes the local SQLite database
├── local_config.py       # Environment-specific config (not tracked in git — create your own)
├── Database.db            # SQLite database (not tracked in git — generated locally)
├── Database.sqbpro        # DB Browser for SQLite project file
├── start_server.bat       # Windows launch script
└── templates/
    ├── form.html           # Trip entry form
    ├── dagboek.html         # Diary/schedule + quotations grid
    └── admin.html           # Admin panel (buses, drivers, records, service, analytics)
```

## Database schema (SQLite)

Key tables: `BusName`, `BusTable`, `Drivers`, `Kwotasie`, `ServiceLog`, `ServiceHistory`, `Tanks`, `BestemmingStreek`.

Historical data was originally migrated from a Microsoft Access database via one-off import scripts (not part of the running app) — Access is used only as a read source for backfilling; SQLite is the permanent source of truth going forward.

## Setup

1. **Clone the repo** and install dependencies:
   ```bash
   pip install fastapi uvicorn jinja2 openpyxl
   ```
2. **Create `local_config.py`** in the project root (this file is gitignored) with any environment-specific paths your setup needs.
3. **Initialize the database:**
   ```bash
   python init_db.py
   ```
4. **Run the server:**
   ```bash
   uvicorn main:app --reload
   ```
   or, on Windows, double-click `start_server.bat`.
5. Open `http://localhost:8000` in your browser.

## Notes

- `Database.db` and `local_config.py` are intentionally excluded from version control since they contain local/client-specific data and environment paths.
- The app is designed to run on a local network (dev laptop → server PC), with Git used to deploy updates.
