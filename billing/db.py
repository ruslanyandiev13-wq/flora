"""
Схема и запросы биллинга (ТЗ "Биллинг / История", раздел 3). Новые таблицы,
не связанные с бизнес-логикой голландского/импортного модулей - используют
ту же базу data/app.db (через db.get_conn()), но не трогают ничего из
существующих таблиц db.py.

Мультитенантность не нужна (компания одна) - org_id жёстко зафиксирован как
DEFAULT_ORG_ID, без экрана переключения организаций. Поле org_id в таблицах
всё же оставлено - дешёвая страховка на случай, если сервис в будущем станет
мультитенантным (см. ТЗ, раздел 6, п.3).
"""
import datetime

import db
import timeutil
from billing.pricing import MIN_TOPUP_TOKENS, MONTHLY_TOKEN_PACKAGE

DEFAULT_ORG_ID = "default"


def init_db():
    conn = db.get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS org_balance (
            org_id TEXT PRIMARY KEY,
            current_balance INTEGER NOT NULL,
            plan_renewal_date TEXT NOT NULL,
            monthly_token_grant INTEGER NOT NULL,
            storage_used_mb REAL DEFAULT 0
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS token_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id TEXT NOT NULL,
            datetime TEXT NOT NULL,
            delta INTEGER NOT NULL,
            reason TEXT NOT NULL,
            related_invoice_id INTEGER,
            balance_after INTEGER NOT NULL,
            FOREIGN KEY (org_id) REFERENCES org_balance(org_id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id TEXT NOT NULL,
            module TEXT NOT NULL,
            uploaded_by TEXT NOT NULL,
            upload_datetime TEXT NOT NULL,
            supplier_template TEXT,
            page_count INTEGER,
            line_items_count INTEGER,
            awb_count INTEGER DEFAULT 0,
            mix_rule_applied INTEGER DEFAULT 0,
            complexity_score INTEGER NOT NULL,
            complexity_label TEXT NOT NULL,
            status TEXT NOT NULL,
            result_file_path TEXT
        )
    """)
    conn.commit()

    # Миграция: колонки для хранилища документов (нужны чат-правкам - см. ТЗ
    # "Чат-помощник и обратная связь", раздел 1). CREATE TABLE IF NOT EXISTS
    # не добавляет колонки в уже существующую (на проде) таблицу invoices,
    # поэтому добавляем явно, пропуская уже применённые миграции.
    existing_cols = {row["name"] for row in c.execute("PRAGMA table_info(invoices)").fetchall()}
    for col, decl in [("source_file_path", "TEXT"), ("parsed_data_json", "TEXT"),
                       ("generation_context_json", "TEXT"), ("source_filename", "TEXT")]:
        if col not in existing_cols:
            c.execute(f"ALTER TABLE invoices ADD COLUMN {col} {decl}")
    conn.commit()

    _migrate_timestamps_to_gmt3(conn, c)

    c.execute("SELECT COUNT(*) FROM org_balance WHERE org_id=?", (DEFAULT_ORG_ID,))
    if c.fetchone()[0] == 0:
        renewal = (timeutil.today() + datetime.timedelta(days=30)).isoformat()
        c.execute("""
            INSERT INTO org_balance (org_id, current_balance, plan_renewal_date, monthly_token_grant)
            VALUES (?, ?, ?, ?)
        """, (DEFAULT_ORG_ID, MONTHLY_TOKEN_PACKAGE, renewal, MONTHLY_TOKEN_PACKAGE))
    conn.commit()
    conn.close()


def _table_exists(c, name):
    return c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                     (name,)).fetchone() is not None


def _migrate_timestamps_to_gmt3(conn, c):
    """Одноразовый сдвиг уже накопленных отметок времени в GMT+3.

    До этой правки время писалось через datetime.now() - то есть по часовому
    поясу сервера (на проде systemd работает в UTC), а users.created_at - и
    вовсе дефолтом SQLite datetime('now'), всегда UTC. Чтобы старые строки в
    истории не расходились с новыми, сдвигаем их один раз:
      - invoices / token_ledger: на (3 - смещение сервера) часов;
      - users.created_at: всегда на +3 часа (там гарантированно UTC).
    Факт применения отмечаем в служебной таблице - повторный запуск ничего
    не делает.
    """
    c.execute("CREATE TABLE IF NOT EXISTS billing_migrations (name TEXT PRIMARY KEY, applied_at TEXT)")
    conn.commit()
    if c.execute("SELECT 1 FROM billing_migrations WHERE name='timestamps_gmt3'").fetchone():
        return

    server_offset_h = (datetime.datetime.now().astimezone().utcoffset()
                       or datetime.timedelta(0)).total_seconds() / 3600.0
    shift_h = 3 - server_offset_h
    # COALESCE - страховка: strftime() вернёт NULL на строке, которую не смог
    # разобрать, и без него такая отметка времени просто пропала бы.
    if shift_h:
        shift = f"{shift_h:+g} hours"
        c.execute("UPDATE invoices SET upload_datetime = COALESCE("
                  "strftime('%Y-%m-%dT%H:%M:%S', upload_datetime, ?), upload_datetime)", (shift,))
        c.execute("UPDATE token_ledger SET datetime = COALESCE("
                  "strftime('%Y-%m-%dT%H:%M:%S', datetime, ?), datetime)", (shift,))
        # Модуль поддержки (ошибки/отзывы/чат) пишет время так же, как биллинг,
        # и показывается в тех же админских таблицах - сдвигаем заодно.
        for table, col in [("processing_errors", "occurred_at"), ("processing_errors", "resolved_at"),
                           ("feedback", "submitted_at"), ("chat_sessions", "started_at"),
                           ("chat_messages", "sent_at"), ("corrections", "created_at"),
                           ("corrections", "applied_at")]:
            if _table_exists(c, table):
                c.execute(f"UPDATE {table} SET {col} = COALESCE("
                          f"strftime('%Y-%m-%dT%H:%M:%S', {col}, ?), {col})", (shift,))
    if _table_exists(c, "users"):
        c.execute("UPDATE users SET created_at = COALESCE("
                  "strftime('%Y-%m-%dT%H:%M:%S', created_at, '+3 hours'), created_at)")
    c.execute("INSERT INTO billing_migrations (name, applied_at) VALUES ('timestamps_gmt3', ?)",
              (timeutil.now_str(),))
    conn.commit()


def get_balance(org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    row = conn.execute("SELECT * FROM org_balance WHERE org_id=?", (org_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def _record_ledger(conn, org_id, delta, reason, balance_after, related_invoice_id=None):
    conn.execute("""
        INSERT INTO token_ledger (org_id, datetime, delta, reason, related_invoice_id, balance_after)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (org_id, timeutil.now_str(), delta, reason,
          related_invoice_id, balance_after))


def charge_tokens(cost, related_invoice_id, org_id=DEFAULT_ORG_ID, reason="invoice_processing"):
    """Списывает cost токенов с баланса организации. Вызывается ТОЛЬКО после
    успешной генерации итогового файла (см. ТЗ, раздел 1: при ошибке разбора
    токены не списываются). reason='chat_correction' - для правок через чат
    (см. ТЗ "Чат-помощник и обратная связь", раздел 2.2)."""
    conn = db.get_conn()
    row = conn.execute("SELECT current_balance FROM org_balance WHERE org_id=?", (org_id,)).fetchone()
    new_balance = row["current_balance"] - cost
    conn.execute("UPDATE org_balance SET current_balance=? WHERE org_id=?", (new_balance, org_id))
    _record_ledger(conn, org_id, -cost, reason, new_balance, related_invoice_id)
    conn.commit()
    conn.close()
    return new_balance


def renew_subscription(org_id=DEFAULT_ORG_ID):
    """Продление подписки: остаток НЕ сгорает, к нему прибавляется пакет
    тарифа (см. ТЗ, раздел 6, п.1), дата продления сдвигается на 30 дней."""
    conn = db.get_conn()
    row = conn.execute("SELECT current_balance, monthly_token_grant FROM org_balance WHERE org_id=?",
                        (org_id,)).fetchone()
    new_balance = row["current_balance"] + row["monthly_token_grant"]
    new_renewal = (timeutil.today() + datetime.timedelta(days=30)).isoformat()
    conn.execute("UPDATE org_balance SET current_balance=?, plan_renewal_date=? WHERE org_id=?",
                 (new_balance, new_renewal, org_id))
    _record_ledger(conn, org_id, row["monthly_token_grant"], "subscription_renewal", new_balance)
    conn.commit()
    conn.close()
    return new_balance


def manual_topup(tokens, comment, org_id=DEFAULT_ORG_ID):
    """Ручное пополнение баланса администратором (счёт, без онлайн-оплаты).
    Минимум MIN_TOPUP_TOKENS токенов - проверяется и на роуте тоже, не только
    здесь (см. ТЗ, раздел 6, п.2)."""
    if tokens < MIN_TOPUP_TOKENS:
        raise ValueError(f"Минимальный объём докупки - {MIN_TOPUP_TOKENS} токенов")
    conn = db.get_conn()
    row = conn.execute("SELECT current_balance FROM org_balance WHERE org_id=?", (org_id,)).fetchone()
    new_balance = row["current_balance"] + tokens
    conn.execute("UPDATE org_balance SET current_balance=? WHERE org_id=?", (new_balance, org_id))
    _record_ledger(conn, org_id, tokens, "manual_adjustment", new_balance)
    conn.commit()

    # Причина/комментарий пополнения - отдельной колонки в token_ledger под
    # это не заведено в ТЗ, поэтому дописываем его прямо в reason, чтобы не
    # потерять (иначе комментарий администратора нигде не сохранится).
    conn.execute("""
        UPDATE token_ledger SET reason = ? WHERE id = (SELECT MAX(id) FROM token_ledger WHERE org_id=?)
    """, (f"manual_adjustment: {comment}", org_id))
    conn.commit()
    conn.close()
    return new_balance


def record_invoice(*, org_id=DEFAULT_ORG_ID, module, uploaded_by, supplier_template,
                    page_count, line_items_count, awb_count, mix_rule_applied,
                    complexity_score, complexity_label, status, result_file_path=None,
                    source_filename=None):
    conn = db.get_conn()
    cur = conn.execute("""
        INSERT INTO invoices (org_id, module, uploaded_by, upload_datetime, supplier_template,
            page_count, line_items_count, awb_count, mix_rule_applied, complexity_score,
            complexity_label, status, result_file_path, source_filename)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (org_id, module, uploaded_by, timeutil.now_str(),
          supplier_template, page_count, line_items_count, awb_count, int(mix_rule_applied),
          complexity_score, complexity_label, status, result_file_path, source_filename))
    conn.commit()
    invoice_id = cur.lastrowid
    conn.close()
    return invoice_id


def attach_document(invoice_id, source_file_path=None, parsed_data=None, generation_context=None):
    """Прикрепляет к уже созданной записи invoices постоянные ссылку на
    исходник и снэпшоты данных, нужные для правок через чат (см. ТЗ
    "Чат-помощник и обратная связь", раздел 1: propose_correction/
    apply_correction работают со структурированными данными документа, не
    с сырым файлом)."""
    import json
    conn = db.get_conn()
    conn.execute("""
        UPDATE invoices SET source_file_path=?, parsed_data_json=?, generation_context_json=?
        WHERE id=?
    """, (source_file_path,
          json.dumps(parsed_data, ensure_ascii=False, default=str) if parsed_data is not None else None,
          json.dumps(generation_context, ensure_ascii=False, default=str) if generation_context is not None else None,
          invoice_id))
    conn.commit()
    conn.close()


def get_invoice(invoice_id, org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    row = conn.execute("SELECT * FROM invoices WHERE id=? AND org_id=?", (invoice_id, org_id)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_invoice_after_correction(invoice_id, result_file_path, generation_context,
                                     complexity_score, complexity_label, line_items_count,
                                     parsed_data=None):
    """Обновляет накладную на месте после 'тяжёлой' правки через чат (ТЗ
    "Чат-помощник и обратная связь", раздел 2.2) - в отличие от обычной
    обработки, здесь НЕ создаётся новая запись в invoices: пересчитанные
    показатели пишутся в ТУ ЖЕ строку, чтобы token_ledger.related_invoice_id
    указывал на исходный документ ("накладная X, повторно списано Y
    токенов"), а не на новую запись."""
    import json
    conn = db.get_conn()
    conn.execute("""
        UPDATE invoices SET result_file_path=?, generation_context_json=?,
            complexity_score=?, complexity_label=?, line_items_count=?
            {}
        WHERE id=?
    """.format(", parsed_data_json=?" if parsed_data is not None else ""),
        ([result_file_path, json.dumps(generation_context, ensure_ascii=False, default=str),
          complexity_score, complexity_label, line_items_count] +
         ([json.dumps(parsed_data, ensure_ascii=False, default=str)] if parsed_data is not None else []) +
         [invoice_id]))
    conn.commit()
    conn.close()


def update_invoice_result(invoice_id, result_file_path, generation_context=None):
    import json
    conn = db.get_conn()
    if generation_context is not None:
        conn.execute("UPDATE invoices SET result_file_path=?, generation_context_json=? WHERE id=?",
                     (result_file_path, json.dumps(generation_context, ensure_ascii=False, default=str),
                      invoice_id))
    else:
        conn.execute("UPDATE invoices SET result_file_path=? WHERE id=?", (result_file_path, invoice_id))
    conn.commit()
    conn.close()


def get_history(org_id=DEFAULT_ORG_ID, module=None, status=None, uploaded_by=None,
                 complexity_label=None, date_from=None, date_to=None, limit=200):
    conn = db.get_conn()
    query = "SELECT * FROM invoices WHERE org_id=?"
    params = [org_id]
    if module:
        query += " AND module=?"
        params.append(module)
    if status:
        query += " AND status=?"
        params.append(status)
    if uploaded_by:
        query += " AND uploaded_by=?"
        params.append(uploaded_by)
    if complexity_label:
        query += " AND complexity_label=?"
        params.append(complexity_label)
    if date_from:
        query += " AND upload_datetime >= ?"
        params.append(date_from)
    if date_to:
        query += " AND upload_datetime <= ?"
        params.append(date_to + "T23:59:59")
    query += " ORDER BY upload_datetime DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_distinct_uploaders(org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    rows = conn.execute("SELECT DISTINCT uploaded_by FROM invoices WHERE org_id=? ORDER BY uploaded_by",
                         (org_id,)).fetchall()
    conn.close()
    return [r["uploaded_by"] for r in rows]


def get_ledger(org_id=DEFAULT_ORG_ID, limit=500):
    conn = db.get_conn()
    rows = conn.execute("SELECT * FROM token_ledger WHERE org_id=? ORDER BY datetime DESC LIMIT ?",
                         (org_id, limit)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_topup_history(org_id=DEFAULT_ORG_ID, limit=200):
    """Только пополнения (delta > 0) - продление подписки, ручная докупка.
    В отличие от get_ledger() (полный журнал, включая списания на
    обработку) - эта выборка предназначена для показа и admin, и user."""
    conn = db.get_conn()
    rows = conn.execute("""
        SELECT * FROM token_ledger WHERE org_id=? AND delta > 0
        ORDER BY datetime DESC LIMIT ?
    """, (org_id, limit)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_storage_used_mb(org_id=DEFAULT_ORG_ID):
    """Информационный показатель (ТЗ, раздел 5, п.4) - без лимита и
    тарификации. Считаем по факту, суммируя размер файлов-результатов,
    которые ещё существуют на диске (временные каталоги периодически
    чистятся, так что часть путей из истории уже не существует)."""
    import os as _os
    conn = db.get_conn()
    rows = conn.execute(
        "SELECT result_file_path FROM invoices WHERE org_id=? AND result_file_path IS NOT NULL",
        (org_id,)).fetchall()
    conn.close()
    total_bytes = 0
    for r in rows:
        try:
            total_bytes += _os.path.getsize(r["result_file_path"])
        except OSError:
            pass
    return round(total_bytes / (1024 * 1024), 2)


def get_daily_spend(org_id=DEFAULT_ORG_ID, since_date=None):
    """Расход (только списания на обработку) по дням, для графика burn-down."""
    conn = db.get_conn()
    query = """
        SELECT substr(datetime, 1, 10) AS day, SUM(-delta) AS spent
        FROM token_ledger
        WHERE org_id=? AND reason='invoice_processing'
    """
    params = [org_id]
    if since_date:
        query += " AND datetime >= ?"
        params.append(since_date)
    query += " GROUP BY day ORDER BY day"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]
