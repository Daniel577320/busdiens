import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent / "Database.db"

conn = sqlite3.connect(DB_PATH)

# Create BusName table (links bus names to their seater category)
conn.execute("""
    CREATE TABLE IF NOT EXISTS BusName (
        ID      INTEGER PRIMARY KEY AUTOINCREMENT,
        BusName TEXT NOT NULL UNIQUE,
        Seater  TEXT NOT NULL
    )
""")

# Create Drivers table
conn.execute("""
    CREATE TABLE IF NOT EXISTS Drivers (
        ID         INTEGER PRIMARY KEY AUTOINCREMENT,
        DriverName TEXT NOT NULL UNIQUE
    )
""")

# Seed BusName from existing Bus records so nothing breaks
conn.execute("""
    INSERT OR IGNORE INTO BusName (BusName, Seater)
    SELECT DISTINCT BusName, '35-Seater' FROM Bus
""")

# Seed Drivers from existing Bus records
conn.execute("""
    INSERT OR IGNORE INTO Drivers (DriverName)
    SELECT DISTINCT Driver FROM Bus
""")

conn.commit()
conn.close()
print("Database initialised successfully.")
