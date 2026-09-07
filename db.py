import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "app.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS tara_mapping (
            fust_code TEXT PRIMARY KEY,
            label TEXT,
            capacity_per_pallet REAL NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS recipients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            label TEXT NOT NULL,
            debnr TEXT NOT NULL,
            naam TEXT NOT NULL
        )
    """)
    conn.commit()

    # начальные значения (можно менять через /dictionaries)
    c.execute("SELECT COUNT(*) FROM settings WHERE key='pallet_cost_usd'")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO settings(key, value) VALUES ('pallet_cost_usd', '1650')")

    c.execute("SELECT COUNT(*) FROM recipients")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO recipients(label, debnr, naam) VALUES (?,?,?)",
                   ("IRIS / Flora Cargo (Вологда)", "IRIS", 'LLC "Flora Cargo"'))

    conn.commit()
    conn.close()


# --- тара -------------------------------------------------------------

def get_tara_mapping():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM tara_mapping ORDER BY fust_code").fetchall()
    conn.close()
    return {r["fust_code"]: dict(r) for r in rows}


def upsert_tara(fust_code, label, capacity_per_pallet):
    conn = get_conn()
    conn.execute("""
        INSERT INTO tara_mapping (fust_code, label, capacity_per_pallet)
        VALUES (?, ?, ?)
        ON CONFLICT(fust_code) DO UPDATE SET label=excluded.label,
            capacity_per_pallet=excluded.capacity_per_pallet
    """, (fust_code, label, capacity_per_pallet))
    conn.commit()
    conn.close()


def delete_tara(fust_code):
    conn = get_conn()
    conn.execute("DELETE FROM tara_mapping WHERE fust_code=?", (fust_code,))
    conn.commit()
    conn.close()


# --- настройки ----------------------------------------------------------

def get_setting(key, default=None):
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_conn()
    conn.execute("""
        INSERT INTO settings(key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (key, str(value)))
    conn.commit()
    conn.close()


# --- получатели / маршруты ----------------------------------------------

def get_recipients():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM recipients ORDER BY label").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_recipient(label, debnr, naam):
    conn = get_conn()
    conn.execute("INSERT INTO recipients(label, debnr, naam) VALUES (?,?,?)", (label, debnr, naam))
    conn.commit()
    conn.close()


def delete_recipient(recipient_id):
    conn = get_conn()
    conn.execute("DELETE FROM recipients WHERE id=?", (recipient_id,))
    conn.commit()
    conn.close()
