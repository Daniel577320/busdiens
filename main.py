from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import sqlite3
from datetime import date
from pathlib import Path

app = FastAPI()

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

DB_PATH = Path(__file__).parent / "Database.db"


# =========================
# DATABASE CONNECTION
# =========================
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn



# =========================
# GOOD/BAD HELPER
# =========================
def calc_good_bad(bus_name: str, kml, conn) -> str:
    """Compare kml against bus Target_KML. Falls back to 'No Target Set' if none defined."""
    if kml is None:
        return "No Target Set"
    row = conn.execute(
        "SELECT Target_KML FROM BusName WHERE BusNames = ?", (bus_name,)
    ).fetchone()
    target = row["Target_KML"] if row and row["Target_KML"] is not None else None
    if target is None:
        return "No Target Set"
    return "Good" if kml >= target else "Bad"

# =========================
# HOME PAGE
# =========================
@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="form.html"
    )
# =========================
# GET ALL BUS Seaters
# =========================

@app.get("/seaters")
def get_seaters():
    conn = get_db()
    seaters = conn.execute("""
        SELECT DISTINCT Seater
        FROM BusName
        ORDER BY Seater
    """).fetchall()
    conn.close()
    return [s["Seater"] for s in seaters]

# =========================
# GET ALL BUS NAMES
# =========================
@app.get("/buses")
def get_buses():
    conn = get_db()
    buses = conn.execute("""
        SELECT DISTINCT BusName
        FROM BusTable
        ORDER BY BusName
    """).fetchall()
    conn.close()
    return [bus["BusName"] for bus in buses]


# =========================
# GET ALL DRIVERS
# =========================
@app.get("/drivers")
def get_drivers():
    conn = get_db()
    drivers = conn.execute("""
        SELECT DriverName
        FROM Drivers
        ORDER BY DriverName
    """).fetchall()
    conn.close()
    return [d["DriverName"] for d in drivers]

# =========================
# GET BUSES BY SEATER
# =========================
@app.get("/buses-by-seater/{seater}")
def get_buses_by_seater(seater: str):
    conn = get_db()
    buses = conn.execute("""
        SELECT BusNames
        FROM BusName
        WHERE Seater = ?
        ORDER BY BusNames
    """, (seater,)).fetchall()
    conn.close()
    return [b["BusNames"] for b in buses]


# =========================
# GET LATEST BUS DATA
# =========================
@app.get("/bus/{bus_name}")
def get_bus_latest(bus_name: str):
    conn = get_db()
    bus = conn.execute("""
        SELECT *
        FROM BusTable
        WHERE BusName = ?
        ORDER BY ID DESC
        LIMIT 1
    """, (bus_name,)).fetchone()
    conn.close()

    if bus:
        return {
            "busName":         bus["BusName"],
            "driverName":      bus["Driver"],
            "currentOdometer": bus["E_Odometer"],
            "fuelUsed":        bus["Fuel_Used"],
            "date":            bus["Date"]
        }

    return JSONResponse(
        content={"error": "Bus not found"},
        status_code=404
    )


# =========================
# GET ALL RECORDS FOR BUS
# =========================
@app.get("/bus/{bus_name}/records")
def get_bus_records(bus_name: str):
    conn = get_db()
    rows = conn.execute("""
        SELECT *
        FROM BusTable
        WHERE BusName = ?
        ORDER BY ID ASC
    """, (bus_name,)).fetchall()
    conn.close()

    # Calculate per-driver averages across all rows
    driver_km_totals = {}
    driver_fuel_totals = {}
    for row in rows:
        driver = row["Driver"]
        km = row["E_Odometer"] - row["B_Odometer"]
        fuel = row["Fuel_Used"]
        if driver not in driver_km_totals:
            driver_km_totals[driver] = 0
            driver_fuel_totals[driver] = 0
        driver_km_totals[driver] += km
        driver_fuel_totals[driver] += fuel

    driver_avg = {}
    for driver in driver_km_totals:
        total_fuel = driver_fuel_totals[driver]
        driver_avg[driver] = (
            round(driver_km_totals[driver] / total_fuel, 2)
            if total_fuel else None
        )

    # Calculate overall bus average across all rows
    total_bus_km   = sum(row["E_Odometer"] - row["B_Odometer"] for row in rows)
    total_bus_fuel = sum(row["Fuel_Used"] for row in rows)
    bus_avg = round(total_bus_km / total_bus_fuel, 2) if total_bus_fuel else None

    results = []
    for row in rows:
        km = row["E_Odometer"] - row["B_Odometer"]
        fuel = row["Fuel_Used"]
        kml = row["KM/L"]
        performance = row["Good/Bad"]
        driver = row["Driver"]

        results.append({
            "date":             row["Date"],
            "driverName":       driver,
            "previousOdometer": row["B_Odometer"],
            "currentOdometer":  row["E_Odometer"],
            "totalKm":          km,
            "fuelUsed":         fuel,
            "kmPerLitre":       kml,
            "busAvg":           bus_avg,
            "performance":      performance,
            "driverAvg":        driver_avg.get(driver)
        })

    return results




# =========================
# GET RECORDS FOR BUS + DRIVER
# =========================
@app.get("/bus/{bus_name}/driver/{driver_name}/records")
def get_bus_driver_records(bus_name: str, driver_name: str):
    conn = get_db()
    rows = conn.execute("""
        SELECT *
        FROM BusTable
        WHERE BusName = ? AND Driver = ?
        ORDER BY ID ASC
    """, (bus_name, driver_name)).fetchall()
    conn.close()

    total_km = sum(row["E_Odometer"] - row["B_Odometer"] for row in rows)
    total_fuel = sum(row["Fuel_Used"] for row in rows)
    driver_avg = round(total_km / total_fuel, 2) if total_fuel else None
    bus_avg = driver_avg  # when filtered to one driver, bus avg for that context = driver avg

    results = []
    for row in rows:
        km = row["E_Odometer"] - row["B_Odometer"]
        results.append({
            "date":             row["Date"],
            "driverName":       row["Driver"],
            "previousOdometer": row["B_Odometer"],
            "currentOdometer":  row["E_Odometer"],
            "totalKm":          km,
            "fuelUsed":         row["Fuel_Used"],
            "kmPerLitre":       row["KM/L"],
            "busAvg":           bus_avg,
            "performance":      row["Good/Bad"],
            "driverAvg":        driver_avg
        })

    return results

# =========================
# CHECK IF BUS HAS RECORDS
# =========================
@app.get("/bus/{bus_name}/has-records")
def bus_has_records(bus_name: str):
    conn = get_db()
    row = conn.execute("""
        SELECT COUNT(*) as cnt FROM BusTable WHERE BusName = ?
    """, (bus_name,)).fetchone()
    conn.close()
    return {"hasRecords": row["cnt"] > 0}

# =========================
# SUBMIT NEW ENTRY
# =========================
@app.post("/entry")
async def submit_entry(request: Request):
    body = await request.json()

    bus_name      = body.get("busName")
    driver_name   = body.get("driverName")
    e_odometer    = body.get("currentOdometer")
    fuel_used     = body.get("fuelUsed")
    start_odometer = body.get("startOdometer")  # only needed for first entry

    if not all([bus_name, driver_name, e_odometer, fuel_used]):
        return JSONResponse(
            content={"error": "Missing fields"},
            status_code=400
        )

    conn = get_db()

    # Get the last odometer reading for this bus to use as B_Odometer
    last = conn.execute("""
        SELECT E_Odometer FROM BusTable
        WHERE BusName = ?
        ORDER BY ID DESC
        LIMIT 1
    """, (bus_name,)).fetchone()

    if last:
        b_odometer = last["E_Odometer"]
    elif start_odometer is not None:
        b_odometer = start_odometer
    else:
        conn.close()
        return JSONResponse(
            content={"error": "No previous odometer reading found. Please provide a starting odometer value."},
            status_code=400
        )

    km = e_odometer - b_odometer
    kml = round(km / fuel_used, 2) if fuel_used else None
    good_bad = calc_good_bad(bus_name, kml, conn)

    conn.execute("""
        INSERT INTO BusTable (BusName, Driver, Date, B_Odometer, E_Odometer, Fuel_Used, [KM/L], [Good/Bad])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        bus_name,
        driver_name,
        str(date.today()),
        b_odometer,
        e_odometer,
        fuel_used,
        kml,
        good_bad
    ))

    conn.commit()
    conn.close()

    return {"status": "success"}


# =========================
# BUS TARGET KML
# =========================
@app.get("/admin/bus/{bus_id}/target-kml")
def get_target_kml(bus_id: int):
    conn = get_db()
    row = conn.execute("SELECT Target_KML FROM BusName WHERE ID = ?", (bus_id,)).fetchone()
    conn.close()
    if not row:
        return JSONResponse(content={"error": "Bus not found"}, status_code=404)
    return {"target_kml": row["Target_KML"]}

@app.put("/admin/bus/{bus_id}/target-kml")
async def set_target_kml(bus_id: int, request: Request):
    body = await request.json()
    target = body.get("target_kml")
    # Allow null to clear the target
    if target is not None:
        try:
            target = float(target)
        except (ValueError, TypeError):
            return JSONResponse(content={"error": "target_kml must be a number or null"}, status_code=400)
    conn = get_db()
    conn.execute("UPDATE BusName SET Target_KML = ? WHERE ID = ?", (target, bus_id))
    conn.commit()
    conn.close()
    return {"status": "ok", "target_kml": target}

# =========================
# ADMIN PAGE
# =========================
@app.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request):
    return templates.TemplateResponse(request=request, name="admin.html")

# --- Admin: Seaters ---
@app.get("/admin/seaters")
def admin_seaters():
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT Seater FROM BusName ORDER BY Seater").fetchall()
    conn.close()
    return [r["Seater"] for r in rows]

# --- Admin: All buses (full rows) ---
@app.get("/admin/buses")
def admin_buses():
    conn = get_db()
    rows = conn.execute("SELECT * FROM BusName ORDER BY Seater, BusNames").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/admin/bus")
async def admin_add_bus(request: Request):
    body = await request.json()
    conn = get_db()
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("BusName")').fetchall() if c["name"] != "ID"]
    inserts = {k: v for k, v in body.items() if k in cols and str(v).strip()}
    if not inserts:
        conn.close()
        return JSONResponse(content={"error": "At least one field is required"}, status_code=400)
    col_clause = ", ".join(f'"{k}"' for k in inserts)
    val_clause = ", ".join("?" for _ in inserts)
    conn.execute(f'INSERT INTO BusName ({col_clause}) VALUES ({val_clause})', (*inserts.values(),))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.put("/admin/bus/{bus_id}")
async def admin_update_bus(bus_id: int, request: Request):
    body = await request.json()
    if not body:
        return JSONResponse(content={"error": "No data provided"}, status_code=400)
    conn = get_db()
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("BusName")').fetchall() if c["name"] != "ID"]
    updates = {k: v for k, v in body.items() if k in cols}
    if not updates:
        conn.close()
        return JSONResponse(content={"error": "No valid columns to update"}, status_code=400)
    set_clause = ", ".join(f'"{k}" = ?' for k in updates)
    conn.execute(f'UPDATE BusName SET {set_clause} WHERE ID = ?', (*updates.values(), bus_id))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/admin/bus/{bus_id}")
def admin_delete_bus(bus_id: int):
    conn = get_db()
    conn.execute("DELETE FROM BusName WHERE ID = ?", (bus_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# --- Admin: All drivers (full rows) ---
@app.get("/admin/drivers")
def admin_drivers():
    conn = get_db()
    rows = conn.execute("SELECT * FROM Drivers ORDER BY DriverName").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/admin/driver")
async def admin_add_driver(request: Request):
    body = await request.json()
    conn = get_db()
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("Drivers")').fetchall() if c["name"] != "ID"]
    inserts = {k: v for k, v in body.items() if k in cols and str(v).strip()}
    if not inserts:
        conn.close()
        return JSONResponse(content={"error": "At least one field is required"}, status_code=400)
    col_clause = ", ".join(f'"{k}"' for k in inserts)
    val_clause = ", ".join("?" for _ in inserts)
    conn.execute(f'INSERT INTO Drivers ({col_clause}) VALUES ({val_clause})', (*inserts.values(),))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.put("/admin/driver/{driver_id}")
async def admin_update_driver(driver_id: int, request: Request):
    body = await request.json()
    if not body:
        return JSONResponse(content={"error": "No data provided"}, status_code=400)
    conn = get_db()
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("Drivers")').fetchall() if c["name"] != "ID"]
    updates = {k: v for k, v in body.items() if k in cols}
    if not updates:
        conn.close()
        return JSONResponse(content={"error": "No valid columns to update"}, status_code=400)
    set_clause = ", ".join(f'"{k}" = ?' for k in updates)
    conn.execute(f'UPDATE Drivers SET {set_clause} WHERE ID = ?', (*updates.values(), driver_id))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/admin/driver/{driver_id}")
def admin_delete_driver(driver_id: int):
    conn = get_db()
    conn.execute("DELETE FROM Drivers WHERE ID = ?", (driver_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# --- Admin: Trip records ---
@app.get("/admin/records")
def admin_records(bus: str = None, driver: str = None):
    conn = get_db()
    query = "SELECT * FROM BusTable WHERE 1=1"
    params = []
    if bus:    query += " AND BusName = ?";  params.append(bus)
    if driver: query += " AND Driver = ?";   params.append(driver)
    query += " ORDER BY ID DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.put("/admin/record/{record_id}")
async def admin_update_record(record_id: int, request: Request):
    body = await request.json()
    bus    = body.get("BusName")
    driver = body.get("Driver")
    dt     = body.get("Date")
    b_odo  = body.get("B_Odometer")
    e_odo  = body.get("E_Odometer")
    fuel   = body.get("Fuel_Used")
    if not all([bus, driver, dt, b_odo is not None, e_odo is not None, fuel]):
        return JSONResponse(content={"error": "Missing fields"}, status_code=400)
    km      = float(e_odo) - float(b_odo)
    kml     = round(km / float(fuel), 2) if fuel else None
    conn = get_db()
    good_bad = calc_good_bad(bus, kml, conn)
    conn.execute("""
        UPDATE BusTable
        SET BusName=?, Driver=?, Date=?, B_Odometer=?, E_Odometer=?, Fuel_Used=?, [KM/L]=?, [Good/Bad]=?
        WHERE ID=?
    """, (bus, driver, dt, b_odo, e_odo, fuel, kml, good_bad, record_id))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/admin/record/{record_id}")
def admin_delete_record(record_id: int):
    conn = get_db()
    conn.execute("DELETE FROM BusTable WHERE ID = ?", (record_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# =========================
# ADMIN: SCHEMA MANAGEMENT
# =========================
MANAGED_TABLES = {"BusName": "BusName", "Drivers": "Drivers", "BusTable": "BusTable"}
PROTECTED_COLS = {"BusName": ["ID","Seater","BusNames","Target_KML"],
                  "Drivers": ["ID","DriverName"],
                  "BusTable": ["ID","BusName","Driver","Date","B_Odometer","E_Odometer","Fuel_Used","KM/L","Good/Bad"]}

@app.get("/admin/schema/{table}")
def admin_schema(table: str):
    if table not in MANAGED_TABLES:
        return JSONResponse(content={"error": "Unknown table"}, status_code=400)
    conn = get_db()
    cols = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    conn.close()
    return [{"name": c["name"], "type": c["type"], "protected": c["name"] in PROTECTED_COLS.get(table, [])} for c in cols]

@app.post("/admin/schema/{table}/add-column")
async def admin_add_column(table: str, request: Request):
    if table not in MANAGED_TABLES:
        return JSONResponse(content={"error": "Unknown table"}, status_code=400)
    body = await request.json()
    col_name = (body.get("name") or "").strip().replace(" ", "_")
    col_type = body.get("type", "TEXT").upper()
    if col_type not in ("TEXT", "INTEGER", "REAL"):
        col_type = "TEXT"
    if not col_name:
        return JSONResponse(content={"error": "Column name required"}, status_code=400)
    conn = get_db()
    existing = [c["name"] for c in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
    if col_name in existing:
        conn.close()
        return JSONResponse(content={"error": "Column already exists"}, status_code=400)
    conn.execute(f'ALTER TABLE "{table}" ADD COLUMN "{col_name}" {col_type}')
    conn.commit()
    conn.close()
    return {"status": "ok", "column": col_name}

@app.delete("/admin/schema/{table}/column/{col_name}")
def admin_delete_column(table: str, col_name: str):
    if table not in MANAGED_TABLES:
        return JSONResponse(content={"error": "Unknown table"}, status_code=400)
    protected = PROTECTED_COLS.get(table, [])
    if col_name in protected:
        return JSONResponse(content={"error": f"Cannot delete protected column '{col_name}'"}, status_code=400)
    conn = get_db()
    cols = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    remaining = [c["name"] for c in cols if c["name"] != col_name]
    col_defs  = [c for c in cols if c["name"] != col_name]
    # SQLite doesn't support DROP COLUMN before 3.35 — use recreate pattern
    sqlite_ver = tuple(int(x) for x in conn.execute("SELECT sqlite_version()").fetchone()[0].split("."))
    if sqlite_ver >= (3, 35, 0):
        conn.execute(f'ALTER TABLE "{table}" DROP COLUMN "{col_name}"')
    else:
        # Recreate table without that column
        type_map = {c["name"]: c["type"] for c in cols}
        col_list = ", ".join(f'"{c}" {type_map[c]}' for c in remaining)
        conn.execute(f'CREATE TABLE "__tmp_{table}" ({col_list})')
        sel = ", ".join(f'"{c}"' for c in remaining)
        conn.execute(f'INSERT INTO "__tmp_{table}" SELECT {sel} FROM "{table}"')
        conn.execute(f'DROP TABLE "{table}"')
        conn.execute(f'ALTER TABLE "__tmp_{table}" RENAME TO "{table}"')
    conn.commit()
    conn.close()
    return {"status": "ok"}
