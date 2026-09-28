import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "app.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# Фермы для раздела «Поставки». Коды TESSA - по таблице закупщика
# (2026-09-28, PS2 = Positano), написания - из HAWB Fresh Solutions и
# инвойсов партии VIKA 21-25.09.
DEFAULT_GROWER_ALIASES = [(alias, grower) for grower, aliases in {
    "POSITANO": ["POSITANO", "POSITANO FARMS", "TESSA-P", "TESSA-PM", "TESSA-PT", "TESSA-PS2"],
    "TESSAROSES": ["TESSAROSES", "TESSA-1", "TESSA-3", "TESSA-A"],
    "GROWER": ["TESSA-D"],
    "ECUANROSE": ["ECUANROSE", "ECUANROS", "TESSA-E1", "TESSA-E2"],
    "ARCOFLOR": ["ARCOFLOR", "ARCOFLOR FLORES", "TESSA-F"],
    "TESSA FARMS": ["TESSA FARMS", "TESSA-FM", "TESSA-FT"],
    "PONTE TRESA": ["PONTE TRESA", "INVERSIONES PONTE", "TESSA-R1", "TESSA-R2", "TESSA-R3"],
    "SOLERA": ["SOLERA", "SOLERA FARMS", "TESSA-S"],
    "QUALITY SERVICE": ["QUALITY SERVICE", "QUALISA SERVICE"],
    "MATIZ": ["MATIZ", "MATIZ ROSES", "ANDES BLOSSOMFARMS", "ANDES BLOSSOM"],
    "MONTEROSAS": ["MONTEROSAS", "MONTEROSASLIMITADA"],
    "ROSAS DEL CORAZON": ["ROSAS DEL CORAZON", "ROSASLESANDI"],
    "STAR ROSES": ["STAR ROSES", "EL CAMPANARIO", "EL CAMPANARIO DE SANTA ANITA"],
    "ECOROSES": ["ECOROSES"],
}.items() for alias in aliases]

# Добавлены 2026-09-28 после первой партии на проде: поставщики «Импорта»,
# которых не было в первом наборе, и написание из HAWB "ROSA PRIMA CIA. LTDA.".
# Засеваются отдельной версией - удалённые закупщиком написания первого
# набора при этом не возвращаются.
GROWER_ALIASES_V2 = [(alias, grower) for grower, aliases in {
    "GARDAEXPORT": ["GARDAEXPORT", "GARDA"],
    "FLORSANI": ["FLORSANI"],
    "ROSAPRIMA": ["ROSAPRIMA", "ROSA PRIMA", "ROSAPRIMA INTERNATIONAL"],
    "CERES": ["CERES", "CERESFARMS", "CERES FARMS"],
    "UTOPIA": ["UTOPIA", "UTOPIA FARMS"],
}.items() for alias in aliases]


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
    # Ассортимент бухгалтера ("Голландия (ассортимент).xlsx") - справочник
    # канонических названий сортов. Обработчик приводит к ним то, что напечатано
    # в паклисте: MH Flowers дописывает в название плантацию ("R Tr Summer Dance
    # Water"), а у бухгалтера позиция называется "R Tr Summer Dance". Иногда,
    # наоборот, плантация - часть названия ("R Tr Yellow Babe Flora Ola"), и
    # трогать его нельзя - отличить одно от другого можно только по этому списку.
    c.execute("""
        CREATE TABLE IF NOT EXISTS assortment (
            name TEXT PRIMARY KEY COLLATE NOCASE,
            bunch INTEGER,
            height TEXT,
            weight TEXT,
            comment TEXT
        )
    """)
    # Ручные замены для случаев, которых нет ни в паклисте, ни в ассортименте
    # ("Blushing Bride 4-6" -> "Serruria Blushing Bride").
    c.execute("""
        CREATE TABLE IF NOT EXISTS variety_aliases (
            printed TEXT PRIMARY KEY COLLATE NOCASE,
            canonical TEXT NOT NULL
        )
    """)
    # Приписки плантаций/марок в конце названия, которые бухгалтер убирает
    # ("Iris Blue Magic Decorum" -> "Iris Blue Magic"). Только явный список:
    # по одному паклисту не отличить приписку от части сорта ("Resq Salmon").
    c.execute("""
        CREATE TABLE IF NOT EXISTS grower_suffixes (
            suffix TEXT PRIMARY KEY COLLATE NOCASE
        )
    """)
    # Раздел «Поставки» (метки Москвы): документы копятся по метке по мере
    # поступления - инвойсы ферм и HAWB форвардера; раскладка коробок по
    # рейсам каждый раз считается заново из них (deliveries.py).
    c.execute("""
        CREATE TABLE IF NOT EXISTS delivery_docs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mark TEXT NOT NULL,
            kind TEXT NOT NULL,          -- 'invoice' | 'hawb'
            doc_key TEXT NOT NULL,       -- номер инвойса / HAWB: повторная загрузка заменяет
            filename TEXT,
            template TEXT,
            data TEXT NOT NULL,          -- JSON разобранного документа
            uploaded_at TEXT,
            uploaded_by TEXT,
            UNIQUE(mark, kind, doc_key)
        )
    """)
    # Ручная привязка коробки к рейсу, когда автоматика ошиблась или
    # вариантов несколько. box_key = "<id документа>:<номер коробки>".
    c.execute("""
        CREATE TABLE IF NOT EXISTS delivery_pins (
            box_key TEXT PRIMARY KEY,
            hawb_doc_id INTEGER NOT NULL
        )
    """)
    # Как называть ферму: коды TESSA (TESSA-P) и написания из HAWB
    # (INVERSIONES PONTE) -> одно имя фермы.
    c.execute("""
        CREATE TABLE IF NOT EXISTS grower_aliases (
            alias TEXT PRIMARY KEY COLLATE NOCASE,
            grower TEXT NOT NULL
        )
    """)
    conn.commit()

    # начальные значения (можно менять через /dictionaries)
    c.execute("SELECT COUNT(*) FROM settings WHERE key='pallet_cost_usd'")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO settings(key, value) VALUES ('pallet_cost_usd', '1650')")

    # Разбор правок бухгалтера 30.09 и 13.09. Засеваем один раз, чтобы
    # удалённые через «Справочники» приписки не возвращались при рестарте.
    c.execute("SELECT COUNT(*) FROM settings WHERE key='grower_suffixes_seeded'")
    if c.fetchone()[0] == 0:
        c.executemany("INSERT OR IGNORE INTO grower_suffixes(suffix) VALUES (?)",
                      [("Decorum",), ("Location Aalsmeer",), ("Water",)])
        c.execute("INSERT INTO settings(key, value) VALUES ('grower_suffixes_seeded', '1')")

    c.execute("SELECT COUNT(*) FROM settings WHERE key='grower_aliases_seeded'")
    if c.fetchone()[0] == 0:
        c.executemany("INSERT OR IGNORE INTO grower_aliases(alias, grower) VALUES (?,?)",
                      DEFAULT_GROWER_ALIASES)
        c.execute("INSERT INTO settings(key, value) VALUES ('grower_aliases_seeded', '1')")
    c.execute("SELECT COUNT(*) FROM settings WHERE key='grower_aliases_seeded_v2'")
    if c.fetchone()[0] == 0:
        c.executemany("INSERT OR IGNORE INTO grower_aliases(alias, grower) VALUES (?,?)",
                      GROWER_ALIASES_V2)
        c.execute("INSERT INTO settings(key, value) VALUES ('grower_aliases_seeded_v2', '1')")
    # Ставка перевозки за кг для раздела «Поставки» (закупщик, 2026-09-28).
    # Дней в пути от вылета до доставки на склад - для «ожидается dd.mm».
    for key, value in (("delivery_rate_kg_ecuador", "8.1"), ("delivery_rate_kg_colombia", "8"),
                       ("delivery_transit_days", "1")):
        c.execute("INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)", (key, value))

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


# --- ассортимент и замены названий --------------------------------------

def get_assortment_names():
    """Множество канонических названий сортов (как их пишет бухгалтер)."""
    conn = get_conn()
    rows = conn.execute("SELECT name FROM assortment").fetchall()
    conn.close()
    return [r["name"] for r in rows]


def assortment_count():
    conn = get_conn()
    n = conn.execute("SELECT COUNT(*) FROM assortment").fetchone()[0]
    conn.close()
    return n


def replace_assortment(rows):
    """rows: [{"name":.., "bunch":.., "height":.., "weight":.., "comment":..}].
    Справочник заменяется целиком - бухгалтер заливает свой файл ассортимента
    как есть, а не правит позиции по одной."""
    conn = get_conn()
    conn.execute("DELETE FROM assortment")
    conn.executemany(
        "INSERT OR REPLACE INTO assortment(name, bunch, height, weight, comment) VALUES (?,?,?,?,?)",
        [(r["name"], r.get("bunch"), r.get("height"), r.get("weight"), r.get("comment"))
         for r in rows])
    conn.commit()
    conn.close()
    return len(rows)


def get_variety_aliases():
    conn = get_conn()
    rows = conn.execute("SELECT printed, canonical FROM variety_aliases ORDER BY printed").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def upsert_variety_alias(printed, canonical):
    conn = get_conn()
    conn.execute("""
        INSERT INTO variety_aliases(printed, canonical) VALUES (?,?)
        ON CONFLICT(printed) DO UPDATE SET canonical=excluded.canonical
    """, (printed.strip(), canonical.strip()))
    conn.commit()
    conn.close()


def delete_variety_alias(printed):
    conn = get_conn()
    conn.execute("DELETE FROM variety_aliases WHERE printed=?", (printed,))
    conn.commit()
    conn.close()


def get_grower_suffixes():
    conn = get_conn()
    rows = conn.execute("SELECT suffix FROM grower_suffixes ORDER BY suffix").fetchall()
    conn.close()
    return [r["suffix"] for r in rows]


def add_grower_suffix(suffix):
    conn = get_conn()
    conn.execute("INSERT OR IGNORE INTO grower_suffixes(suffix) VALUES (?)", (" ".join(suffix.split()),))
    conn.commit()
    conn.close()


def delete_grower_suffix(suffix):
    conn = get_conn()
    conn.execute("DELETE FROM grower_suffixes WHERE suffix=?", (suffix,))
    conn.commit()
    conn.close()


# --- Поставки ----------------------------------------------------------------

def get_grower_aliases():
    conn = get_conn()
    rows = conn.execute("SELECT alias, grower FROM grower_aliases ORDER BY grower, alias").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def upsert_grower_alias(alias, grower):
    conn = get_conn()
    conn.execute("""
        INSERT INTO grower_aliases(alias, grower) VALUES (?, ?)
        ON CONFLICT(alias) DO UPDATE SET grower=excluded.grower
    """, (" ".join(alias.split()).upper(), " ".join(grower.split()).upper()))
    conn.commit()
    conn.close()


def delete_grower_alias(alias):
    conn = get_conn()
    conn.execute("DELETE FROM grower_aliases WHERE alias=?", (alias,))
    conn.commit()
    conn.close()


def save_delivery_doc(mark, kind, doc_key, filename, template, data_json, uploaded_at, uploaded_by):
    """Повторная загрузка того же документа (метка + вид + номер) заменяет
    старую версию, id сохраняется - ручные привязки коробок не теряются."""
    conn = get_conn()
    conn.execute("""
        INSERT INTO delivery_docs(mark, kind, doc_key, filename, template, data, uploaded_at, uploaded_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(mark, kind, doc_key) DO UPDATE SET
            filename=excluded.filename, template=excluded.template, data=excluded.data,
            uploaded_at=excluded.uploaded_at, uploaded_by=excluded.uploaded_by
    """, (mark, kind, doc_key, filename, template, data_json, uploaded_at, uploaded_by))
    conn.commit()
    conn.close()


def get_delivery_docs(mark=None):
    conn = get_conn()
    if mark is None:
        rows = conn.execute("SELECT * FROM delivery_docs ORDER BY id").fetchall()
    else:
        rows = conn.execute("SELECT * FROM delivery_docs WHERE mark=? ORDER BY id", (mark,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delivery_marks():
    conn = get_conn()
    rows = conn.execute("""
        SELECT mark, SUM(kind='invoice') AS invoices, SUM(kind='hawb') AS hawbs,
               MAX(uploaded_at) AS last_upload
        FROM delivery_docs GROUP BY mark ORDER BY MAX(uploaded_at) DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_delivery_doc(doc_id):
    conn = get_conn()
    conn.execute("DELETE FROM delivery_docs WHERE id=?", (doc_id,))
    conn.execute("DELETE FROM delivery_pins WHERE box_key LIKE ? OR hawb_doc_id=?",
                 (f"{doc_id}:%", doc_id))
    conn.commit()
    conn.close()


def get_delivery_pins():
    conn = get_conn()
    rows = conn.execute("SELECT box_key, hawb_doc_id FROM delivery_pins").fetchall()
    conn.close()
    return {r["box_key"]: r["hawb_doc_id"] for r in rows}


def set_delivery_pin(box_key, hawb_doc_id):
    conn = get_conn()
    if hawb_doc_id is None:
        conn.execute("DELETE FROM delivery_pins WHERE box_key=?", (box_key,))
    else:
        conn.execute("""
            INSERT INTO delivery_pins(box_key, hawb_doc_id) VALUES (?, ?)
            ON CONFLICT(box_key) DO UPDATE SET hawb_doc_id=excluded.hawb_doc_id
        """, (box_key, hawb_doc_id))
    conn.commit()
    conn.close()
