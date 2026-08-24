from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import sqlite3
from datetime import date, timedelta
from pathlib import Path
from fastapi.responses import FileResponse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tempfile, calendar
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

app = FastAPI()

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

DB_PATH = Path(__file__).parent / "Database.db"

# Every tank shares the same maximum capacity for now (litres)
TANK_MAX_CAPACITY = 2000


# =========================
# DATABASE CONNECTION
# =========================
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_service_log_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName TEXT NOT NULL,
            LastServiced DATE,
            Odometer_At_Service REAL,
            Comments TEXT,
            RecordType TEXT NOT NULL DEFAULT 'Service',
            ServiceType TEXT
        )
    """)
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("ServiceLog")').fetchall()]
    if "RecordType" not in cols:
        conn.execute("ALTER TABLE ServiceLog ADD COLUMN RecordType TEXT NOT NULL DEFAULT 'Service'")
    if "ServiceType" not in cols:
        conn.execute("ALTER TABLE ServiceLog ADD COLUMN ServiceType TEXT")


def ensure_bus_table_has_tank_column(conn):
    """Add a TankName column to BusTable if it doesn't already exist (safe to call repeatedly)."""
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("BusTable")').fetchall()]
    if "TankName" not in cols:
        conn.execute('ALTER TABLE BusTable ADD COLUMN TankName TEXT')



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
# GET TRIP TYPES (Kwotasie.tipeTrip options)
# =========================
# Fixed short list rather than a DB lookup table — matches the handful of
# trip types the client actually quotes against. Add more here if the
# client starts using additional trip types.
TRIP_TYPES = ["Dag-Retoer", "Toer-Retoer", "Toer-Oorslaap"]


@app.get("/trip-types")
def get_trip_types():
    return TRIP_TYPES

# =========================
# GET ALL BUS NAMES
# =========================
@app.get("/buses")
def get_buses():
    conn = get_db()
    buses = conn.execute("""
        SELECT DISTINCT BusNames AS BusName
        FROM BusName
        WHERE TRIM(BusNames) <> ''
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
# GET ALL TANKS
# =========================
@app.get("/tanks")
def get_tanks():
    conn = get_db()
    tanks = conn.execute("""
        SELECT TankName
        FROM Tanks
        ORDER BY TankName
    """).fetchall()
    conn.close()
    return [t["TankName"] for t in tanks]


# =========================
# REFUEL A TANK (or all tanks)
# =========================
@app.post("/tanks/refuel")
async def refuel_tank(request: Request):
    body = await request.json()

    tank_name     = body.get("tankName")
    fuel_amount   = body.get("fuelAmount")
    fill_selected = bool(body.get("fillSelected"))
    fill_all      = bool(body.get("fillAll"))

    conn = get_db()
    today = str(date.today())

    # Fill every tank to full — takes priority over everything else
    if fill_all:
        conn.execute(
            "UPDATE Tanks SET Est_Capacity = ?, Date_Refuelled = ?",
            (TANK_MAX_CAPACITY, today)
        )
        conn.commit()
        conn.close()
        return {"status": "success", "filledAll": True, "level": TANK_MAX_CAPACITY}

    if not tank_name:
        conn.close()
        return JSONResponse(content={"error": "tankName is required"}, status_code=400)

    tank_row = conn.execute(
        "SELECT Tank_ID FROM Tanks WHERE TankName = ?", (tank_name,)
    ).fetchone()
    if not tank_row:
        conn.close()
        return JSONResponse(content={"error": f"Tank '{tank_name}' not found"}, status_code=400)

    # Fill just the selected tank to full
    if fill_selected:
        conn.execute(
            "UPDATE Tanks SET Est_Capacity = ?, Date_Refuelled = ? WHERE TankName = ?",
            (TANK_MAX_CAPACITY, today, tank_name)
        )
        conn.commit()
        conn.close()
        return {"status": "success", "tankName": tank_name, "level": TANK_MAX_CAPACITY}

    # Otherwise add the entered amount, capped at the tank's max capacity
    if fuel_amount is None:
        conn.close()
        return JSONResponse(content={"error": "fuelAmount is required"}, status_code=400)
    try:
        fuel_amount = float(fuel_amount)
    except (ValueError, TypeError):
        conn.close()
        return JSONResponse(content={"error": "fuelAmount must be a number"}, status_code=400)

    conn.execute("""
        UPDATE Tanks
        SET Est_Capacity = MIN(?, COALESCE(Est_Capacity, 0) + ?), Date_Refuelled = ?
        WHERE TankName = ?
    """, (TANK_MAX_CAPACITY, fuel_amount, today, tank_name))
    conn.commit()

    new_level = conn.execute(
        "SELECT Est_Capacity FROM Tanks WHERE TankName = ?", (tank_name,)
    ).fetchone()["Est_Capacity"]
    conn.close()
    return {"status": "success", "tankName": tank_name, "level": new_level}


# =========================
# ADMIN: TANK LEVELS
# =========================
@app.get("/admin/tanks")
def admin_get_tanks():
    conn = get_db()
    rows = conn.execute("""
        SELECT Tank_ID, TankName, Est_Capacity, Date_Refuelled
        FROM Tanks
        ORDER BY TankName
    """).fetchall()
    conn.close()
    return [
        {
            "tankId": r["Tank_ID"],
            "tankName": r["TankName"],
            "currentFuel": r["Est_Capacity"] if r["Est_Capacity"] is not None else 0,
            "maxCapacity": TANK_MAX_CAPACITY,
            "dateRefuelled": r["Date_Refuelled"],
        }
        for r in rows
    ]


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
    tank_name     = body.get("tankName")
    start_odometer = body.get("startOdometer")  # only needed for first entry

    if not all([bus_name, driver_name, e_odometer, fuel_used, tank_name]):
        return JSONResponse(
            content={"error": "Missing fields"},
            status_code=400
        )

    conn = get_db()

    # Make sure the selected tank actually exists before we touch anything
    tank_row = conn.execute(
        "SELECT Tank_ID, Est_Capacity FROM Tanks WHERE TankName = ?", (tank_name,)
    ).fetchone()
    if not tank_row:
        conn.close()
        return JSONResponse(
            content={"error": f"Tank '{tank_name}' not found"},
            status_code=400
        )

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
    if km < 0:
        conn.close()
        return JSONResponse(
            content={"error": f"Current odometer ({e_odometer}) can't be less than the previous odometer ({b_odometer}). This would result in a negative distance."},
            status_code=400
        )
    kml = round(km / fuel_used, 2) if fuel_used else None
    good_bad = calc_good_bad(bus_name, kml, conn)

    ensure_bus_table_has_tank_column(conn)

    conn.execute("""
        INSERT INTO BusTable (BusName, Driver, Date, B_Odometer, E_Odometer, Fuel_Used, [KM/L], [Good/Bad], TankName)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        bus_name,
        driver_name,
        str(date.today()),
        b_odometer,
        e_odometer,
        fuel_used,
        kml,
        good_bad,
        tank_name
    ))

    # Deduct the fuel used for this trip from the selected tank's remaining capacity
    conn.execute("""
        UPDATE Tanks
        SET Est_Capacity = MAX(0, COALESCE(Est_Capacity, 0) - ?)
        WHERE TankName = ?
    """, (fuel_used, tank_name))

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


@app.get("/dagboek", response_class=HTMLResponse)
def dagboek_page(request: Request):
    return templates.TemplateResponse(request=request, name="dagboek.html")


# =========================
# DAGBOEK / KWOTASIE
# =========================
# This is the app's own source of truth going forward, stored in the same
# Database.db every other part of the app already uses — plain sqlite3, no
# external drivers. Bus/driver are stored as plain names here (not foreign
# keys), matching how the dropdowns on the Dagboek page already work against
# /buses and /drivers.
#
# Historical quotes from the client's old Access database were brought in
# once via import_from_access.py (see that file) — this table is not synced
# back to Access, and nothing here talks to Access at runtime.

def ensure_kwotasie_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS Kwotasie (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            datum TEXT,
            bus TEXT,
            klient TEXT,
            bestemming TEXT,
            tydOplaai TEXT,
            tydTerug TEXT,
            bedrag REAL,
            bestuurder TEXT,
            opmerkings TEXT,
                aanvaar INTEGER NOT NULL DEFAULT 0,
                betaal INTEGER NOT NULL DEFAULT 0,
            access_id INTEGER,
            finaliseerDatum TEXT,
            seater TEXT,
            dateCreated TEXT,
            tipeTrip TEXT,
            datumTerug TEXT,
            tripGroupId INTEGER,
            trailer INTEGER NOT NULL DEFAULT 0,
            faktuur TEXT
        )
    """)
    # Existing databases created before these columns existed won't have them
    # yet — add whichever are missing (safe to call repeatedly).
    cols = [c["name"] for c in conn.execute('PRAGMA table_info("Kwotasie")').fetchall()]
    for new_col in ("betaal", "finaliseerDatum", "seater", "dateCreated", "tipeTrip", "datumTerug", "tripGroupId", "trailer", "faktuur"):
        if new_col not in cols:
            conn.execute(f'ALTER TABLE Kwotasie ADD COLUMN {new_col} TEXT')


def ensure_bestemming_streek_table(conn):
    """Local lookup of destination -> region (Streek), backfilled one-time
    from the Access Bestemming table via import_streek_from_access.py.
    A trip's bestemming column can list several comma-separated destination
    names, so this is resolved per-destination and re-joined at read time —
    same approach already used for multi-destination bestemming itself."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS BestemmingStreek (
            bestemmingNaam TEXT PRIMARY KEY,
            streek TEXT
        )
    """)


def _resolve_streek(conn, bestemming: str) -> str:
    if not bestemming:
        return ""
    names = [n.strip() for n in bestemming.split(",") if n.strip()]
    if not names:
        return ""
    placeholders = ",".join("?" for _ in names)
    rows = conn.execute(
        f"SELECT streek FROM BestemmingStreek WHERE bestemmingNaam IN ({placeholders})",
        names,
    ).fetchall()
    streke = sorted({r["streek"] for r in rows if r["streek"]})
    return ", ".join(streke)


@app.get("/dagboek/history")
def dagboek_history_entries():
    conn = get_db()
    ensure_kwotasie_table(conn)
    ensure_bestemming_streek_table(conn)
    rows = conn.execute("""
        SELECT id, datum, bus, klient AS klientID, bestemming, tydOplaai, tydTerug,
               bedrag, bestuurder, opmerkings, finaliseerDatum, seater, dateCreated, tipeTrip,
               datumTerug, tripGroupId, aanvaar, betaal, trailer, faktuur
        FROM Kwotasie
        ORDER BY datum DESC, id DESC
    """).fetchall()
    results = []
    for r in rows:
        d = dict(r)
        d["streek"] = _resolve_streek(conn, d["bestemming"])
        results.append(d)
    conn.close()
    return results


@app.get("/dagboek/entries")
def dagboek_list_entries(year: int, month: int, type: str):
    aanvaar = 1 if type == "dagboek" else 0
    month_str = f"{year:04d}-{month:02d}"
    conn = get_db()
    ensure_kwotasie_table(conn)
    rows = conn.execute("""
        SELECT id, datum, bus, klient AS klientID, bestemming, tydOplaai, tydTerug,
               bedrag, bestuurder, opmerkings, finaliseerDatum, seater, dateCreated, tipeTrip,
               datumTerug, tripGroupId, aanvaar, betaal, trailer, faktuur
        FROM Kwotasie
        WHERE aanvaar = ? AND substr(datum, 1, 7) = ?
        ORDER BY datum, id
    """, (aanvaar, month_str)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


@app.post("/dagboek/entries")
async def dagboek_create_entry(request: Request):
    body = await request.json()
    aanvaar = 1 if body.get("type") == "dagboek" else 0
    conn = get_db()
    ensure_kwotasie_table(conn)
    cur = conn.execute("""
        INSERT INTO Kwotasie (datum, bus, klient, bestemming, tydOplaai, tydTerug,
                               bedrag, bestuurder, opmerkings, aanvaar,
                               seater, tipeTrip, dateCreated, datumTerug, betaal,
                               trailer, faktuur)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        body.get("datum") or str(date.today()),
        body.get("bus") or "",
        body.get("klientID") or "",
        body.get("bestemming") or "",
        body.get("tydOplaai") or "",
        body.get("tydTerug") or "",
        _parse_float(body.get("bedrag")),
        body.get("bestuurder") or "",
        body.get("opmerkings") or "",
        aanvaar,
        body.get("seater") or "",
        body.get("tipeTrip") or "",
        str(date.today()),  # dateCreated is always stamped now, never client-supplied
        body.get("datumTerug") or "",
        1 if str(body.get("betaal") or "").strip().lower() in {"1", "true", "yes", "checked", "on"} else 0,
        1 if str(body.get("trailer") or "").strip().lower() in {"1", "true", "yes", "checked", "on"} else 0,
        body.get("faktuur") or "",
    ))
    conn.commit()
    new_id = cur.lastrowid
    # A freshly-created row is its own trip group until (if ever) it gets
    # fanned out into multiple days on finalisation — see dagboek_accept_entry.
    conn.execute("UPDATE Kwotasie SET tripGroupId = ? WHERE id = ?", (new_id, new_id))
    conn.commit()
    conn.close()
    return {"status": "success", "id": new_id}


# Map the grid's field names (what the frontend sends) onto actual Kwotasie
# columns. klientID -> klient because the grid/dropdowns refer to the client
# by the name "klientID" even though it's a free-text business name, not a
# numeric id — keeping that mapping here means we don't have to touch every
# call site in dagboek.html.
_DAGBOEK_FIELDS = {
    "datum": "datum",
    "bus": "bus",
    "klientID": "klient",
    "bestemming": "bestemming",
    "tydOplaai": "tydOplaai",
    "tydTerug": "tydTerug",
    "bedrag": "bedrag",
    "bestuurder": "bestuurder",
    "opmerkings": "opmerkings",
    "finaliseerDatum": "finaliseerDatum",
    "seater": "seater",
    "tipeTrip": "tipeTrip",
    "datumTerug": "datumTerug",
    "betaal": "betaal",
    "trailer": "trailer",
    "faktuur": "faktuur",
}


@app.patch("/dagboek/entries/{trip_id}")
async def dagboek_update_entry(trip_id: int, request: Request):
    body = await request.json()
    field = body.get("field")
    value = body.get("value")
    column = _DAGBOEK_FIELDS.get(field)
    if column is None:
        return JSONResponse(content={"error": f"Unknown field: {field}"}, status_code=400)
    if field == "bedrag":
        value = _parse_float(value)
    elif field in ("betaal", "trailer"):
        value = 1 if str(value).strip().lower() in {"1", "true", "yes", "checked", "on"} else 0
    conn = get_db()
    ensure_kwotasie_table(conn)
    conn.execute(f"UPDATE Kwotasie SET {column} = ? WHERE id = ?", (value, trip_id))
    conn.commit()
    conn.close()
    return {"status": "success"}


def _date_range_inclusive(start_str: str, end_str: str):
    """List of ISO date strings from start to end inclusive. Falls back to
    just [start] if end is missing/unparseable/before start, so a malformed
    Datum Terug never blocks finalising a quote."""
    try:
        start = date.fromisoformat(start_str)
        end = date.fromisoformat(end_str)
    except (TypeError, ValueError):
        return [start_str]
    if end < start:
        return [start_str]
    days = []
    current = start
    while current <= end:
        days.append(current.isoformat())
        current += timedelta(days=1)
    return days


@app.post("/dagboek/entries/{trip_id}/accept")
def dagboek_accept_entry(trip_id: int):
    conn = get_db()
    ensure_kwotasie_table(conn)
    row = conn.execute("SELECT * FROM Kwotasie WHERE id = ?", (trip_id,)).fetchone()
    if row is None:
        conn.close()
        return JSONResponse(content={"error": "Not found"}, status_code=404)

    trip_type = (row["tipeTrip"] or "").strip()
    departure = row["datum"]
    terug = row["datumTerug"] or row["datum"]

    if trip_type == "Toer-Oorslaap":
        # One Dagboek entry for every day the bus is away, departure through
        # return inclusive.
        dates = _date_range_inclusive(departure, terug)
    elif trip_type == "Toer-Retoer":
        # Exactly two entries: the day the trip leaves and the day it's back.
        dates = sorted({departure, terug or departure})
    else:
        # Dag-Retoer and anything else: single same-day trip, unchanged.
        dates = [departure]

    group_id = trip_id  # links every generated day back to the original quote

    # The first day reuses the original row in place, so its id/access_id and
    # any history references to it stay intact.
    conn.execute("""
        UPDATE Kwotasie
        SET aanvaar = 1, datum = ?, tripGroupId = ?
        WHERE id = ?
    """, (dates[0], group_id, trip_id))

    # Extra days (return day of a Toer-Retoer, or every day in between for a
    # Toer-Oorslaap) become their own cloned Dagboek rows. Only the first day
    # keeps the quoted amount — the rest are logging the bus being away, not
    # additional billed amounts, so they don't inflate the monthly total.
    for d in dates[1:]:
        conn.execute("""
            INSERT INTO Kwotasie (datum, bus, klient, bestemming, tydOplaai, tydTerug,
                                   bedrag, bestuurder, opmerkings, aanvaar,
                                   seater, tipeTrip, dateCreated, finaliseerDatum,
                                   datumTerug, tripGroupId, trailer, faktuur)
            SELECT ?, bus, klient, bestemming, tydOplaai, tydTerug,
                   NULL, bestuurder, opmerkings, 1,
                   seater, tipeTrip, dateCreated, finaliseerDatum,
                   datumTerug, ?, trailer, faktuur
            FROM Kwotasie WHERE id = ?
        """, (d, group_id, trip_id))

    conn.commit()
    conn.close()
    return {"status": "success", "days_created": len(dates)}


@app.delete("/dagboek/entries/{trip_id}")
def dagboek_delete_entry(trip_id: int):
    conn = get_db()
    ensure_kwotasie_table(conn)
    conn.execute("DELETE FROM Kwotasie WHERE id = ?", (trip_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}


def _parse_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


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
    if km < 0:
        return JSONResponse(content={"error": "Current odometer can't be less than the previous odometer (negative distance)."}, status_code=400)
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
    ensure_bus_table_has_tank_column(conn)
    row = conn.execute(
        "SELECT TankName, Fuel_Used FROM BusTable WHERE ID = ?", (record_id,)
    ).fetchone()
    if not row:
        conn.close()
        return JSONResponse(content={"error": "Record not found"}, status_code=404)

    conn.execute("DELETE FROM BusTable WHERE ID = ?", (record_id,))

    # Refund the fuel this trip used back into the tank it came from, capped
    # at the tank's max capacity. (Distance-to-next-service for the bus is
    # derived live from its latest remaining odometer reading, so removing
    # this trip already restores that distance automatically — no separate
    # bookkeeping needed there.)
    tank_name = row["TankName"]
    fuel_used = row["Fuel_Used"]
    if tank_name and fuel_used:
        conn.execute("""
            UPDATE Tanks
            SET Est_Capacity = MIN(?, COALESCE(Est_Capacity, 0) + ?)
            WHERE TankName = ?
        """, (TANK_MAX_CAPACITY, fuel_used, tank_name))

    conn.commit()
    conn.close()
    return {"status": "ok"}

# =========================
# ADMIN: SCHEMA MANAGEMENT
# =========================
MANAGED_TABLES = {"BusName": "BusName", "Drivers": "Drivers", "BusTable": "BusTable"}
PROTECTED_COLS = {"BusName": ["ID","Seater","BusNames","Target_KML","Service_Interval_Km"],
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

# =========================
# ANALYTICS PAGE
# =========================
@app.get("/analytics/service-cards")
def analytics_service_cards():
    """Return the status of every service type for every bus — one card per
    (bus, service type) pair — resetting that type's interval after a matching
    service is logged."""
    service_schedules = (
        ("Normal", 15000, ("Normal", "Routine Service", "General Service")),
        ("Gear Box Oil", 50000, ("Gear Box Oil",)),
        ("Antifreeze", 80000, ("Antifreeze",)),
    )
    conn = get_db()
    ensure_service_log_table(conn)
    buses = conn.execute("""
        SELECT bn.BusNames,
               (
                    SELECT bt.E_Odometer
                    FROM BusTable bt
                    WHERE bt.BusName = bn.BusNames
                    ORDER BY bt.ID DESC
                    LIMIT 1
                ) AS current_odometer
        FROM BusName bn
        ORDER BY bn.BusNames
    """).fetchall()

    cards = []
    for bus in buses:
        current_odometer = bus["current_odometer"]

        if current_odometer is None:
            for service_type, interval, service_type_aliases in service_schedules:
                cards.append({
                    "busName": bus["BusNames"],
                    "distanceToService": None,
                    "serviceType": service_type,
                    "category": "Fine",
                })
            continue

        for service_type, interval, service_type_aliases in service_schedules:
            placeholders = ", ".join("?" for _ in service_type_aliases)
            params = [bus["BusNames"], *service_type_aliases]
            last_service = conn.execute(f"""
                SELECT Odometer_At_Service
                FROM ServiceLog
                WHERE BusName = ?
                  AND RecordType = 'Service'
                  AND ServiceType IN ({placeholders})
                ORDER BY ID DESC
                LIMIT 1
            """, params).fetchone()

            baseline = (last_service["Odometer_At_Service"]
                        if last_service and last_service["Odometer_At_Service"] is not None
                        else 0)
            distance_since_service = float(current_odometer) - float(baseline)

            if distance_since_service <= 0:
                # Just serviced (or the reading hasn't moved past the service
                # point yet) — the full interval is still remaining.
                distance = interval
            else:
                remainder = distance_since_service % interval
                # remainder == 0 here means a whole number of intervals has
                # passed since the service, i.e. it's due right now.
                distance = 0 if remainder == 0 else interval - remainder

            if distance < 1000:
                category = "Urgent"
            elif distance <= 5000:
                category = "Upcoming"
            else:
                category = "Fine"

            cards.append({
                "busName": bus["BusNames"],
                "distanceToService": round(distance, 1),
                "serviceType": service_type,
                "category": category,
            })

    conn.close()
    return cards

@app.get("/analytics", response_class=HTMLResponse)
def analytics_page(request: Request):
    return templates.TemplateResponse(request=request, name="analytics.html")

# --- seater list for filter dropdowns ---
@app.get("/analytics/seaters")
def analytics_seaters():
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT Seater FROM BusName ORDER BY Seater").fetchall()
    conn.close()
    return [r["Seater"] for r in rows]

# --- fleet overview ---
@app.get("/analytics/fleet-overview")
def analytics_fleet_overview(period: str = "month"):
    from datetime import date, timedelta
    today = date.today()
    since = (today - timedelta(days=7)).isoformat() if period == "week" else today.replace(day=1).isoformat()
    conn = get_db()
    rows = conn.execute("""
        SELECT bt.*, (bt.E_Odometer - bt.B_Odometer) AS km
        FROM BusTable bt WHERE bt.Date >= ?
    """, (since,)).fetchall()
    conn.close()
    if not rows:
        return {"total_km":0,"total_fuel":0,"avg_kml":None,"best_driver":None,
                "worst_driver":None,"best_bus":None,"total_entries":0,"period":period,
                "best_driver_avg":None,"worst_driver_avg":None,"best_bus_avg":None}
    total_km   = sum(r["km"] for r in rows)
    total_fuel = sum(r["Fuel_Used"] for r in rows)
    avg_kml    = round(total_km / total_fuel, 2) if total_fuel else None
    drv_km, drv_fuel = {}, {}
    bus_km, bus_fuel = {}, {}
    for r in rows:
        d, b = r["Driver"], r["BusName"]
        drv_km[d]   = drv_km.get(d,0)   + r["km"];  drv_fuel[d]  = drv_fuel.get(d,0)  + r["Fuel_Used"]
        bus_km[b]   = bus_km.get(b,0)   + r["km"];  bus_fuel[b]  = bus_fuel.get(b,0)  + r["Fuel_Used"]
    drv_avg = {d: round(drv_km[d]/drv_fuel[d],2) for d in drv_km if drv_fuel[d]}
    bus_avg = {b: round(bus_km[b]/bus_fuel[b],2) for b in bus_km if bus_fuel[b]}
    best_driver  = max(drv_avg, key=drv_avg.get) if drv_avg else None
    worst_driver = min(drv_avg, key=drv_avg.get) if drv_avg else None
    best_bus     = max(bus_avg, key=bus_avg.get)  if bus_avg  else None
    return {"total_km":total_km,"total_fuel":total_fuel,"avg_kml":avg_kml,
            "best_driver":best_driver,"worst_driver":worst_driver,"best_bus":best_bus,
            "total_entries":len(rows),"period":period,
            "best_driver_avg":drv_avg.get(best_driver),"worst_driver_avg":drv_avg.get(worst_driver),
            "best_bus_avg":bus_avg.get(best_bus)}

# --- bus performance (optional seater filter) ---
@app.get("/analytics/bus-performance")
def analytics_bus_performance(seater: str = None):
    conn = get_db()
    if seater:
        rows = conn.execute("""
            SELECT bt.BusName, bt.Driver, bt.Date,
                   (bt.E_Odometer - bt.B_Odometer) AS km, bt.Fuel_Used
            FROM BusTable bt
            JOIN BusName bn ON bt.BusName = bn.BusNames
            WHERE bn.Seater = ?
            ORDER BY bt.Date ASC
        """, (seater,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT bt.BusName, bt.Driver, bt.Date,
                   (bt.E_Odometer - bt.B_Odometer) AS km, bt.Fuel_Used
            FROM BusTable bt
            ORDER BY bt.Date ASC
        """).fetchall()
    conn.close()
    bus_km, bus_fuel = {}, {}
    date_km, date_fuel = {}, {}
    bd_km, bd_fuel, bd_trips = {}, {}, {}
    for r in rows:
        b, d, dt = r["BusName"], r["Driver"], r["Date"]
        bus_km[b]  = bus_km.get(b,0)  + r["km"];  bus_fuel[b]  = bus_fuel.get(b,0)  + r["Fuel_Used"]
        date_km[dt] = date_km.get(dt,0) + r["km"]; date_fuel[dt] = date_fuel.get(dt,0) + r["Fuel_Used"]
        key = (b, d)
        bd_km[key]   = bd_km.get(key,0)   + r["km"]
        bd_fuel[key] = bd_fuel.get(key,0) + r["Fuel_Used"]
        bd_trips[key]= bd_trips.get(key,0)+ 1
    avg_kml_per_bus = sorted(
        [{"bus":b,"avg_kml":round(bus_km[b]/bus_fuel[b],2)} for b in bus_km if bus_fuel[b]],
        key=lambda x: x["avg_kml"], reverse=True)
    efficiency_over_time = [
        {"date":dt,"avg_kml":round(date_km[dt]/date_fuel[dt],2)}
        for dt in sorted(date_km) if date_fuel[dt]]
    bus_driver_breakdown = sorted(
        [{"bus":k[0],"driver":k[1],
          "avg_kml":round(bd_km[k]/bd_fuel[k],2),
          "total_km":bd_km[k],"trips":bd_trips[k]}
         for k in bd_km if bd_fuel[k]],
        key=lambda x:(x["bus"],-x["avg_kml"]))
    return {"avg_kml_per_bus":avg_kml_per_bus,
            "efficiency_over_time":efficiency_over_time,
            "bus_driver_breakdown":bus_driver_breakdown}

# --- driver performance (optional seater filter) ---
@app.get("/analytics/driver-performance")
def analytics_driver_performance(seater: str = None):
    conn = get_db()
    if seater:
        rows = conn.execute("""
            SELECT bt.Driver, bt.Date,
                   (bt.E_Odometer - bt.B_Odometer) AS km, bt.Fuel_Used
            FROM BusTable bt
            JOIN BusName bn ON bt.BusName = bn.BusNames
            WHERE bn.Seater = ?
            ORDER BY bt.Date ASC
        """, (seater,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT Driver, Date,
                   (E_Odometer - B_Odometer) AS km, Fuel_Used
            FROM BusTable ORDER BY Date ASC
        """).fetchall()
    conn.close()
    drv_km, drv_fuel, drv_trips, drv_dates = {}, {}, {}, {}
    for r in rows:
        d = r["Driver"]
        drv_km[d]    = drv_km.get(d,0)    + r["km"]
        drv_fuel[d]  = drv_fuel.get(d,0)  + r["Fuel_Used"]
        drv_trips[d] = drv_trips.get(d,0) + 1
        if d not in drv_dates: drv_dates[d] = []
        drv_dates[d].append({"date":r["Date"],
            "kml":round(r["km"]/r["Fuel_Used"],2) if r["Fuel_Used"] else None})
    overall_avg = {d:round(drv_km[d]/drv_fuel[d],2) for d in drv_km if drv_fuel[d]}
    fleet_avg   = round(sum(drv_km.values())/sum(drv_fuel.values()),2) if sum(drv_fuel.values()) else None
    drivers = sorted([
        {"driver":d,"avg_kml":overall_avg.get(d),
         "total_km":drv_km[d],"trips":drv_trips[d],"trend":drv_dates[d]}
        for d in drv_km], key=lambda x:(x["avg_kml"] or 0), reverse=True)
    return {"drivers":drivers,"fleet_avg":fleet_avg}


# =========================
# SERVICE MANAGEMENT
# =========================

@app.get("/admin/service-interval/{bus_id}")
def get_service_interval(bus_id: int):
    conn = get_db()
    row = conn.execute("SELECT Service_Interval_Km FROM BusName WHERE ID = ?", (bus_id,)).fetchone()
    conn.close()
    if not row:
        return JSONResponse(content={"error": "Bus not found"}, status_code=404)
    return {"service_interval_km": row["Service_Interval_Km"]}

@app.put("/admin/service-interval/{bus_id}")
async def set_service_interval(bus_id: int, request: Request):
    body = await request.json()
    interval = body.get("service_interval_km")
    if interval is not None:
        try:
            interval = float(interval)
        except (ValueError, TypeError):
            return JSONResponse(content={"error": "service_interval_km must be a number or null"}, status_code=400)
    conn = get_db()
    # Add column if it doesn't exist yet
    existing = [c["name"] for c in conn.execute('PRAGMA table_info("BusName")').fetchall()]
    if "Service_Interval_Km" not in existing:
        conn.execute('ALTER TABLE "BusName" ADD COLUMN "Service_Interval_Km" REAL')
    conn.execute("UPDATE BusName SET Service_Interval_Km = ? WHERE ID = ?", (interval, bus_id))
    conn.commit()
    conn.close()
    return {"status": "ok", "service_interval_km": interval}

@app.get("/analytics/service-status")
def analytics_service_status():
    """Return service status for all buses, sorted by urgency."""
    from datetime import date as dt_date
    conn = get_db()

    # Ensure columns exist
    existing_cols = [c["name"] for c in conn.execute('PRAGMA table_info("BusName")').fetchall()]
    if "Service_Interval_Km" not in existing_cols:
        conn.execute('ALTER TABLE "BusName" ADD COLUMN "Service_Interval_Km" REAL')
        conn.commit()

    existing_service_cols = [c["name"] for c in conn.execute('PRAGMA table_info("ServiceLog")').fetchall()] \
        if conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ServiceLog'").fetchone() else []

    # Ensure ServiceLog table exists
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID               INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName          TEXT NOT NULL,
            LastServiced     DATE,
            Odometer_At_Service REAL,
            Comments         TEXT
        )
    """)
    conn.commit()

    # Get all buses with Service_Interval_Km
    buses = conn.execute("SELECT * FROM BusName ORDER BY BusNames").fetchall()

    # Get latest odometer per bus from BusTable
    odo_rows = conn.execute("""
        SELECT BusName, E_Odometer as odo, Date
        FROM BusTable
        WHERE ID IN (
            SELECT MAX(ID) FROM BusTable GROUP BY BusName
        )
    """).fetchall()
    odo_map = {r["BusName"]: {"odo": r["odo"], "date": r["Date"]} for r in odo_rows}

    # Compute daily-km average per bus using last 60 days of trip records
    from datetime import date as _date, timedelta
    cutoff_60 = (_date.today() - timedelta(days=60)).isoformat()
    recent_rows = conn.execute("""
        SELECT BusName, SUM(E_Odometer - B_Odometer) as total_km,
               MIN(Date) as first_date, MAX(Date) as last_date,
               COUNT(*) as trips
        FROM BusTable
        WHERE Date >= ? AND (E_Odometer - B_Odometer) > 0
        GROUP BY BusName
    """, (cutoff_60,)).fetchall()
    daily_km_map = {}
    today_iso = _date.today()
    for r in recent_rows:
        try:
            first = _date.fromisoformat(r["first_date"])
            last  = _date.fromisoformat(r["last_date"])
            span  = max((last - first).days, 1)  # avoid div/0; at least 1 day
            # use span between first and last trip date for a realistic average
            daily_km_map[r["BusName"]] = r["total_km"] / span
        except Exception:
            pass

    # Get latest service log per bus
    svc_rows = conn.execute("""
        SELECT BusName, LastServiced, Odometer_At_Service, Comments
        FROM ServiceLog
        WHERE ID IN (
            SELECT MAX(ID) FROM ServiceLog GROUP BY BusName
        )
    """).fetchall()
    svc_map = {r["BusName"]: dict(r) for r in svc_rows}

    conn.close()

    # Status priority order for sorting: Overdue=0, Due Soon=1, Upcoming=2, Normal=3, No Interval=4
    STATUS_ORDER = {"Overdue": 0, "Due Soon": 1, "Upcoming": 2, "Normal": 3, "No Interval": 4}

    results = []
    today = dt_date.today()

    for bus in buses:
        name = bus["BusNames"]
        interval = bus["Service_Interval_Km"]
        current_odo = odo_map.get(name, {}).get("odo")
        svc = svc_map.get(name, {})
        last_serviced = svc.get("LastServiced")
        odo_at_service = svc.get("Odometer_At_Service")
        comments = svc.get("Comments")

        next_service = None
        km_to_service = None
        status = "No Interval"
        est_days = None

        if interval and odo_at_service is not None:
            next_service = odo_at_service + interval
            if current_odo is not None:
                km_to_service = next_service - current_odo
                if km_to_service > 2000:
                    status = "Normal"
                elif km_to_service > 1000:
                    status = "Upcoming"
                elif km_to_service >= 0:
                    status = "Due Soon"
                else:
                    status = "Overdue"

                # Estimate days to service using rolling 60-day daily-km average
                daily_km = daily_km_map.get(name)
                if daily_km and daily_km > 0 and km_to_service is not None:
                    est_days = round(km_to_service / daily_km)
                else:
                    est_days = None
        elif interval and current_odo is not None:
            # Interval set but no service log yet — treat current odo as baseline
            status = "No Service Record"

        results.append({
            "busName": name,
            "seater": bus["Seater"],
            "currentOdometer": current_odo,
            "lastServiced": last_serviced,
            "odometerAtService": odo_at_service,
            "serviceIntervalKm": interval,
            "nextServiceOdometer": round(next_service, 1) if next_service is not None else None,
            "kmToService": round(km_to_service, 1) if km_to_service is not None else None,
            "status": status,
            "estimatedDaysToService": est_days,
            "comments": comments,
        })

    # Sort by urgency
    results.sort(key=lambda x: STATUS_ORDER.get(x["status"], 5))
    return results


@app.post("/service/log")
async def log_service(request: Request):
    """Record a service event for a bus — updates LastServiced, Odometer_At_Service."""
    body = await request.json()
    bus_name  = body.get("busName")
    current_odo = body.get("currentOdometer")
    comments  = body.get("comments", "")

    if not bus_name or current_odo is None:
        return JSONResponse(content={"error": "busName and currentOdometer are required"}, status_code=400)

    from datetime import date as dt_date
    today = str(dt_date.today())

    conn = get_db()
    ensure_service_log_table(conn)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID               INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName          TEXT NOT NULL,
            LastServiced     DATE,
            Odometer_At_Service REAL,
            Comments         TEXT
        )
    """)
    conn.execute("""
        INSERT INTO ServiceLog
            (BusName, LastServiced, Odometer_At_Service, Comments, RecordType, ServiceType)
        VALUES (?, ?, ?, ?, 'Service', 'General Service')
    """, (bus_name, today, float(current_odo), comments))
    conn.commit()
    conn.close()
    return {"status": "ok", "lastServiced": today, "odometerAtService": float(current_odo)}


@app.get("/service/log/{bus_name}")
def get_service_log(bus_name: str):
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID               INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName          TEXT NOT NULL,
            LastServiced     DATE,
            Odometer_At_Service REAL,
            Comments         TEXT
        )
    """)
    rows = conn.execute("""
        SELECT * FROM ServiceLog WHERE BusName = ? ORDER BY ID DESC
    """, (bus_name,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# =========================
# ADMIN: SERVICE HISTORY
# =========================

@app.get("/admin/service-history")
def admin_service_history(bus: str = None, record_type: str = None):
    """Return service and repair history, optionally filtered by bus and type."""
    conn = get_db()
    ensure_service_log_table(conn)
    query = "SELECT * FROM ServiceLog WHERE 1=1"
    params = []
    if bus:
        query += " AND BusName = ?"
        params.append(bus)
    if record_type in ("Service", "Repair"):
        query += " AND RecordType = ?"
        params.append(record_type)
    query += " ORDER BY ID DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


@app.post("/admin/service-history")
async def admin_add_service_history(request: Request):
    body = await request.json()
    record_type = body.get("recordType")
    bus_name = body.get("busName")
    current_odo = body.get("currentOdometer")
    service_type = (body.get("serviceType") or "").strip()
    comments = (body.get("comments") or "").strip()

    if record_type not in ("Service", "Repair"):
        return JSONResponse(content={"error": "recordType must be Service or Repair"}, status_code=400)
    if not bus_name or current_odo is None or not comments:
        return JSONResponse(content={"error": "Bus, current odometer and comments are required"}, status_code=400)
    if record_type == "Service" and not service_type:
        return JSONResponse(content={"error": "Service type is required"}, status_code=400)

    try:
        current_odo = float(current_odo)
    except (ValueError, TypeError):
        return JSONResponse(content={"error": "Current odometer must be a number"}, status_code=400)

    conn = get_db()
    ensure_service_log_table(conn)
    bus_exists = conn.execute("SELECT 1 FROM BusName WHERE BusNames = ?", (bus_name,)).fetchone()
    if not bus_exists:
        conn.close()
        return JSONResponse(content={"error": f"Bus '{bus_name}' not found"}, status_code=400)
    conn.execute("""
        INSERT INTO ServiceLog
            (BusName, LastServiced, Odometer_At_Service, Comments, RecordType, ServiceType)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (bus_name, str(date.today()), current_odo, comments, record_type,
           service_type if record_type == "Service" else "Repair"))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/admin/service-history/{record_id}")
def admin_delete_service_record(record_id: int):
    conn = get_db()
    conn.execute("DELETE FROM ServiceLog WHERE ID = ?", (record_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.post("/service/log")
async def log_service(request: Request):
    """Record a service event for a bus — updates LastServiced, Odometer_At_Service."""
    body = await request.json()
    bus_name  = body.get("busName")
    current_odo = body.get("currentOdometer")
    comments  = body.get("comments", "")

    if not bus_name or current_odo is None:
        return JSONResponse(content={"error": "busName and currentOdometer are required"}, status_code=400)

    from datetime import date as dt_date
    today = str(dt_date.today())

    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID               INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName          TEXT NOT NULL,
            LastServiced     DATE,
            Odometer_At_Service REAL,
            Comments         TEXT
        )
    """)
    conn.execute("""
        INSERT INTO ServiceLog (BusName, LastServiced, Odometer_At_Service, Comments)
        VALUES (?, ?, ?, ?)
    """, (bus_name, today, float(current_odo), comments))
    conn.commit()
    conn.close()
    return {"status": "ok", "lastServiced": today, "odometerAtService": float(current_odo)}


@app.get("/service/log/{bus_name}")
def get_service_log(bus_name: str):
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ServiceLog (
            ID               INTEGER PRIMARY KEY AUTOINCREMENT,
            BusName          TEXT NOT NULL,
            LastServiced     DATE,
            Odometer_At_Service REAL,
            Comments         TEXT
        )
    """)
    rows = conn.execute("""
        SELECT * FROM ServiceLog WHERE BusName = ? ORDER BY ID DESC
    """, (bus_name,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/dagboek/export")
def dagboek_export(year: int, month: int, day: int = 0):
    """
    Export Dagboek entries for a given year/month (and optional day)
    to an Excel file. Opens print-ready: landscape A4, fits to one page.
    """
    from fastapi.responses import FileResponse
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
    import tempfile, calendar

    month_str = f"{year:04d}-{month:02d}"

    # Helper function to format time from HH:MM to HHhMM
    def format_time_for_excel(time_str):
        """Convert HH:MM or HH:MM:SS to HHhMM format (e.g., 08:30 -> 08H30)"""
        if not time_str:
            return ""
        time_str = str(time_str).strip()
        if ':' in time_str:
            parts = time_str.split(':')
            hours = parts[0].zfill(2)
            minutes = parts[1].zfill(2)
            return f"{hours}H{minutes}"
        return time_str

    conn = get_db()
    rows = conn.execute("""
        SELECT id, datum, bus, klient, bestemming, tydOplaai, tydTerug,
            bestuurder, opmerkings, seater, tipeTrip, trailer
        FROM Kwotasie
        WHERE aanvaar = 1 AND substr(datum, 1, 7) = ?
        ORDER BY datum, id
    """, (month_str,)).fetchall()
    bus_rows = conn.execute("SELECT BusNames, Seater FROM BusName").fetchall()
    conn.close()

    seater_map = {}
    for b in bus_rows:
        seater_map[b["BusNames"]] = b["Seater"] or ""

    rows = [dict(r) for r in rows]
    if day:
        day_str = f"{year:04d}-{month:02d}-{day:02d}"
        rows    = [r for r in rows if (r.get("datum") or "").startswith(day_str)]

    for r in rows:
        bus = (r.get("bus") or "").strip()
        r["_export_seater"] = (r.get("seater") or "").strip() or (seater_map.get(bus) or "").strip()
        r["_export_pickup_time"] = format_time_for_excel(r.get("tydOplaai") or "")
        r["_export_return_time"] = format_time_for_excel(r.get("tydTerug") or "")
        r["_export_trailer"] = "Yes" if r.get("trailer") else "No"

    # ── Styles ────────────────────────────────────────────────────────────

    DAY_FILLS = {
        "MAANDAG":  PatternFill("solid", fgColor="4472C4"),  # blue
        "DINSDAG":  PatternFill("solid", fgColor="ED7D31"),  # orange
        "WOENSDAG": PatternFill("solid", fgColor="C33699"),  # purple
        "DONDERDAG":PatternFill("solid", fgColor="4FBF90"),  # teal
        "VRYDAG":   PatternFill("solid", fgColor="8DAE2B"),  # olive
        "SATERDAG": PatternFill("solid", fgColor="D9534F"),  # red
        "SONDAG":   PatternFill("solid", fgColor="27AE60"),  # green
    }
    WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")

    thin   = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Columns: 11 total (A–K) ───────────────────────────────────────────
    # A:NR  B:DEPOT  C:BESTUURDER  D:SEATER  E:BUS  F:TRAILER
    # G:PICK UP  H:BESTEMMING  I:TYD OPLAAI  J:TYD TERUG  K:OPMERKINGS
    LAST_COL     = "K"
    LAST_COL_NUM = 11
    COL_WIDTHS   = [5, 11, 20, 12, 18, 9, 36, 36, 11, 11, 26]
    HEADERS      = ["NR:", "DEPOT:", "DRIVER", "SEATER:", "BUS:", "TRAILER:",
                    "PICK UP", "DESTINATION", "T P/UP", "T BACK", "COMMENTS"]
    MERGE_RANGE  = f"A{{r}}:{LAST_COL}{{r}}"   # e.g. "A1:K1"

    def hdr_cell(ws, row, col, value, size=10):
        c           = ws.cell(row=row, column=col, value=value)
        c.font      = Font(name="Arial", bold=True, size=size)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BORDER
        return c

    def data_cell(ws, row, col, value, bold=False, align="left", alt=False):
        c           = ws.cell(row=row, column=col, value=value)
        c.font      = Font(name="Arial", size=10, bold=bold)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border    = BORDER
        c.fill      = WHITE_FILL
        return c

    # ── Date labels ───────────────────────────────────────────────────────
    month_name  = calendar.month_name[month].upper()
    day_label   = f"{day:02d} {month_name} {year}" if day else f"{month_name} {year}"
    day_of_week = ""
    if day:
        import datetime as dt
        days_af     = ["MAANDAG","DINSDAG","WOENSDAG","DONDERDAG",
                       "VRYDAG","SATERDAG","SONDAG"]
        day_of_week = days_af[dt.date(year, month, day).weekday()]

    # ── Workbook / sheet ──────────────────────────────────────────────────
    wb       = openpyxl.Workbook()
    ws       = wb.active
    ws.title = "BESTUURDERS CLIPBOARD"

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ── Row 1: Title ──────────────────────────────────────────────────────
    ws.merge_cells(MERGE_RANGE.format(r=1))
    c           = ws.cell(row=1, column=1, value="DAY PROGRAM")
    c.font      = Font(name="Arial", bold=True, size=15)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # ── Row 2: Date / day-of-week bar ────────────────────────────────────
    ws.merge_cells(MERGE_RANGE.format(r=2))
    date_str    = f"{day_of_week}, {day_label}" if day_of_week else day_label
    d2          = ws.cell(row=2, column=1, value=date_str)
    day_fill    = DAY_FILLS.get(day_of_week)
    d2.fill     = day_fill
    # Use dark text for light fills (VRYDAG olive / SATERDAG red use white)
    light_days  = {"MAANDAG","DINSDAG","WOENSDAG","DONDERDAG","SATERDAG","SONDAG"}
    txt_color   = "FFFFFF" if day_of_week in light_days else "1F3864"
    d2.font     = Font(name="Arial", bold=True, color=txt_color, size=12)
    d2.alignment= Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 26

    # ── Row 3: Column headers ─────────────────────────────────────────────
    for col, h in enumerate(HEADERS, 1):
        # Smaller font for the two time columns
        hdr_cell(ws, 3, col, h, size=8 if col in (6, 9, 10) else 10)
    ws.row_dimensions[3].height = 26

    # ── Row 4+: Data ──────────────────────────────────────────────────────
    sorted_rows = sorted(rows, key=lambda r: (r.get("_export_pickup_time") or "", r.get("id") or 0))

    for i, r in enumerate(sorted_rows):
        rn  = 4 + i
        alt = (i % 2 == 1)
        bus = (r.get("bus") or "").strip()
        pickup_time = r.get("_export_pickup_time") or ""
        return_time = r.get("_export_return_time") or ""
        ws.row_dimensions[rn].height = 22

        data_cell(ws, rn,  1, i + 1,                      align="center", bold=True, alt=alt)
        data_cell(ws, rn,  2, r.get("_export_pickup_time") or "",   align="center",            alt=alt)  # DEPOT (Pickup time in HHhMM format)
        data_cell(ws, rn,  3, r.get("bestuurder") or "",                             alt=alt)
        data_cell(ws, rn,  4, r.get("_export_seater") or "", align="center",            alt=alt)
        data_cell(ws, rn,  5, bus,                                                   alt=alt)
        data_cell(ws, rn,  6, r.get("_export_trailer") or "",                        alt=alt)  # Trailer: Yes/blank
        data_cell(ws, rn,  7, r.get("klient") or "",                               alt=alt)  # Pick Up
        data_cell(ws, rn,  8, r.get("bestemming") or "",                             alt=alt)
        data_cell(ws, rn,  9, pickup_time,                 align="center",            alt=alt)
        data_cell(ws, rn, 10, return_time,                 align="center",            alt=alt)
        data_cell(ws, rn, 11, r.get("opmerkings") or "",                             alt=alt)

    # Pad to at least 15 rows so the table always looks complete
    filled      = len(sorted_rows)
    target_rows = max(filled, 15)
    for i in range(filled, target_rows):
        rn  = 4 + i
        alt = (i % 2 == 1)
        ws.row_dimensions[rn].height = 18
        data_cell(ws, rn, 1, i + 1, align="center", bold=True, alt=alt)
        for col in range(2, LAST_COL_NUM + 1):
            data_cell(ws, rn, col, "", alt=alt)

    last_data_row = 3 + target_rows

    # ── Print settings ────────────────────────────────────────────────────
    ws.sheet_properties = WorksheetProperties(
        pageSetUpPr=PageSetupProperties(fitToPage=True)
    )
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9   # A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins           = PageMargins(
        left=0.25, right=0.25,
        top=0.4,   bottom=0.4,
        header=0.2, footer=0.2,
    )
    ws.print_options.horizontalCentered = True
    ws.print_area = f"A1:{LAST_COL}{last_data_row}"

    # ── Save & return ─────────────────────────────────────────────────────
    filename = f"DAGPROGRAM_{year}_{month:02d}_{day:02d}.xlsx"
    tmp      = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()

    return FileResponse(
        path=tmp.name,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        filename=filename,
        background=None,
    )
