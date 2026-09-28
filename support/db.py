"""
Схема и запросы для чат-помощника и обратной связи (ТЗ "Чат-помощник и
обратная связь", разделы 3-5). Использует ту же базу data/app.db, не трогает
голландский/импортный модули напрямую.

Таблица `corrections` - НЕ из ТЗ дословно, а моя добавка: ТЗ описывает
propose_correction()/apply_correction(correction_id) как function-calling
инструменты модели, но не даёт схему хранения "предложенной, но ещё не
применённой" правки между двумя вызовами - без таблицы её негде держать
между сообщениями чата. Схема requires минимальный набор полей для diff'а
"было -> станет" плюс тип (light/heavy, см. ТЗ раздел 2.2) для биллинга.
"""
import json

import db
import timeutil

DEFAULT_ORG_ID = "default"


def init_db():
    conn = db.get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS processing_errors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id TEXT NOT NULL,
            module TEXT NOT NULL,
            occurred_at TEXT NOT NULL,
            uploaded_by TEXT,
            source_filename TEXT,
            error_type TEXT,
            error_message TEXT NOT NULL,
            stack_trace TEXT,
            raised_via_chat INTEGER DEFAULT 0,
            chat_session_id TEXT,
            related_document_id INTEGER,
            resolved INTEGER DEFAULT 0,
            resolved_at TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id TEXT NOT NULL,
            document_id INTEGER NOT NULL,
            submitted_by TEXT NOT NULL,
            submitted_at TEXT NOT NULL,
            comment TEXT NOT NULL,
            source TEXT NOT NULL,
            chat_session_id TEXT,
            status TEXT DEFAULT 'new',
            FOREIGN KEY (document_id) REFERENCES invoices(id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_sessions (
            id TEXT PRIMARY KEY,
            org_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            started_at TEXT NOT NULL,
            related_document_id INTEGER
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            sent_at TEXT NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            tool_calls TEXT,
            FOREIGN KEY (session_id) REFERENCES chat_sessions(id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS corrections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            chat_session_id TEXT,
            field_path TEXT NOT NULL,
            current_value TEXT,
            proposed_value TEXT,
            reason TEXT,
            correction_type TEXT NOT NULL,
            created_at TEXT NOT NULL,
            applied INTEGER DEFAULT 0,
            applied_at TEXT,
            FOREIGN KEY (document_id) REFERENCES invoices(id)
        )
    """)
    conn.commit()
    conn.close()


# --- processing_errors ----------------------------------------------------

def log_error(module, error_message, *, org_id=DEFAULT_ORG_ID, uploaded_by=None,
              source_filename=None, error_type="exception", stack_trace=None,
              raised_via_chat=False, chat_session_id=None, related_document_id=None):
    conn = db.get_conn()
    conn.execute("""
        INSERT INTO processing_errors (org_id, module, occurred_at, uploaded_by, source_filename,
            error_type, error_message, stack_trace, raised_via_chat, chat_session_id, related_document_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (org_id, module, timeutil.now_str(), uploaded_by,
          source_filename, error_type, error_message, stack_trace, int(raised_via_chat),
          chat_session_id, related_document_id))
    conn.commit()
    conn.close()


def get_errors(org_id=DEFAULT_ORG_ID, resolved=None, limit=200):
    conn = db.get_conn()
    query = "SELECT * FROM processing_errors WHERE org_id=?"
    params = [org_id]
    if resolved is not None:
        query += " AND resolved=?"
        params.append(int(resolved))
    query += " ORDER BY occurred_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def resolve_error(error_id):
    conn = db.get_conn()
    conn.execute("UPDATE processing_errors SET resolved=1, resolved_at=? WHERE id=?",
                 (timeutil.now_str(), error_id))
    conn.commit()
    conn.close()


# --- feedback ---------------------------------------------------------------

def submit_feedback(document_id, submitted_by, comment, source="feedback_form",
                     chat_session_id=None, org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    conn.execute("""
        INSERT INTO feedback (org_id, document_id, submitted_by, submitted_at, comment, source,
            chat_session_id, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, 'new')
    """, (org_id, document_id, submitted_by, timeutil.now_str(),
          comment, source, chat_session_id))
    conn.commit()
    conn.close()


def get_feedback(org_id=DEFAULT_ORG_ID, module=None, status=None, source=None,
                  date_from=None, date_to=None, submitted_by=None, limit=200):
    conn = db.get_conn()
    query = """
        SELECT feedback.*, invoices.module AS doc_module, invoices.source_filename AS doc_filename
        FROM feedback LEFT JOIN invoices ON invoices.id = feedback.document_id
        WHERE feedback.org_id=?
    """
    params = [org_id]
    if module:
        query += " AND invoices.module=?"
        params.append(module)
    if status:
        query += " AND feedback.status=?"
        params.append(status)
    if source:
        query += " AND feedback.source=?"
        params.append(source)
    if submitted_by:
        query += " AND feedback.submitted_by=?"
        params.append(submitted_by)
    if date_from:
        query += " AND feedback.submitted_at >= ?"
        params.append(date_from)
    if date_to:
        query += " AND feedback.submitted_at <= ?"
        params.append(date_to + "T23:59:59")
    query += " ORDER BY feedback.submitted_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def count_new_feedback(org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    n = conn.execute("SELECT COUNT(*) FROM feedback WHERE org_id=? AND status='new'",
                      (org_id,)).fetchone()[0]
    conn.close()
    return n


def set_feedback_status(feedback_id, status):
    conn = db.get_conn()
    conn.execute("UPDATE feedback SET status=? WHERE id=?", (status, feedback_id))
    conn.commit()
    conn.close()


# --- chat_sessions / chat_messages ------------------------------------------

def create_session(session_id, user_id, related_document_id=None, org_id=DEFAULT_ORG_ID):
    conn = db.get_conn()
    conn.execute("""
        INSERT INTO chat_sessions (id, org_id, user_id, started_at, related_document_id)
        VALUES (?, ?, ?, ?, ?)
    """, (session_id, org_id, user_id, timeutil.now_str(),
          related_document_id))
    conn.commit()
    conn.close()


def get_session(session_id):
    conn = db.get_conn()
    row = conn.execute("SELECT * FROM chat_sessions WHERE id=?", (session_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def add_message(session_id, role, content, tool_calls=None):
    conn = db.get_conn()
    conn.execute("""
        INSERT INTO chat_messages (session_id, sent_at, role, content, tool_calls)
        VALUES (?, ?, ?, ?, ?)
    """, (session_id, timeutil.now_str(), role, content,
          json.dumps(tool_calls, ensure_ascii=False) if tool_calls is not None else None))
    conn.commit()
    conn.close()


def get_messages(session_id):
    conn = db.get_conn()
    rows = conn.execute("SELECT * FROM chat_messages WHERE session_id=? ORDER BY id",
                         (session_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# --- corrections -------------------------------------------------------------

def create_correction(document_id, field_path, current_value, proposed_value, reason,
                       correction_type, chat_session_id=None):
    conn = db.get_conn()
    cur = conn.execute("""
        INSERT INTO corrections (document_id, chat_session_id, field_path, current_value,
            proposed_value, reason, correction_type, created_at, applied)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)
    """, (document_id, chat_session_id, field_path, json.dumps(current_value, ensure_ascii=False, default=str),
          json.dumps(proposed_value, ensure_ascii=False, default=str), reason, correction_type,
          timeutil.now_str()))
    conn.commit()
    correction_id = cur.lastrowid
    conn.close()
    return correction_id


def get_correction(correction_id):
    conn = db.get_conn()
    row = conn.execute("SELECT * FROM corrections WHERE id=?", (correction_id,)).fetchone()
    conn.close()
    if not row:
        return None
    d = dict(row)
    d["current_value"] = json.loads(d["current_value"]) if d["current_value"] is not None else None
    d["proposed_value"] = json.loads(d["proposed_value"]) if d["proposed_value"] is not None else None
    return d


def mark_correction_applied(correction_id):
    conn = db.get_conn()
    conn.execute("UPDATE corrections SET applied=1, applied_at=? WHERE id=?",
                 (timeutil.now_str(), correction_id))
    conn.commit()
    conn.close()
