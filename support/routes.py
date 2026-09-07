"""
Роуты для чат-помощника и обратной связи (ТЗ "Чат-помощник и обратная
связь", разделы 2-5).
"""
from flask import Blueprint, request, render_template, redirect, url_for, flash, jsonify

from auth import login_required, admin_required, current_user
from billing import db as billing_db
import support.db as support_db
import support.chat as chat

support_bp = Blueprint("support", __name__)


# --- чат ---------------------------------------------------------------

@support_bp.route("/chat/recent-documents")
@login_required
def chat_recent_documents():
    rows = billing_db.get_history(uploaded_by=current_user()["username"], limit=5)
    return jsonify([{"id": r["id"], "filename": r["source_filename"],
                       "date": r["upload_datetime"], "module": r["module"]} for r in rows])


@support_bp.route("/chat/start", methods=["POST"])
@login_required
def chat_start():
    document_id = request.json.get("document_id") if request.is_json else request.form.get("document_id")
    session_id = chat.start_session(related_document_id=document_id)
    return jsonify({"session_id": session_id})


@support_bp.route("/chat/message", methods=["POST"])
@login_required
def chat_message():
    payload = request.get_json(silent=True) or {}
    session_id = payload.get("session_id")
    text = (payload.get("text") or "").strip()
    if not session_id or not text:
        return jsonify({"error": "session_id и text обязательны"}), 400
    sess = support_db.get_session(session_id)
    if not sess or sess["user_id"] != current_user()["username"]:
        return jsonify({"error": "Сессия не найдена"}), 404
    try:
        reply = chat.send_message(session_id, text)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({"reply": reply})


# --- обратная связь после обработки -------------------------------------

@support_bp.route("/support/feedback", methods=["POST"])
@login_required
def submit_feedback():
    document_ids = request.form.getlist("document_id")
    comment = (request.form.get("comment") or "").strip()
    next_url = request.form.get("next") or url_for("index")
    if not comment or not document_ids:
        return redirect(next_url)
    username = current_user()["username"]
    for doc_id in document_ids:
        support_db.submit_feedback(int(doc_id), username, comment, source="feedback_form")
    flash("Спасибо, комментарий передан разработчику", "ok")
    return redirect(next_url)


@support_bp.route("/support/my-feedback")
@login_required
def my_feedback():
    rows = support_db.get_feedback(submitted_by=current_user()["username"])
    return render_template("my_feedback.html", rows=rows)


# --- админ: обратная связь и ошибки --------------------------------------

@support_bp.route("/admin/feedback")
@admin_required
def admin_feedback():
    module = request.args.get("module") or None
    status = request.args.get("status") or None
    source = request.args.get("source") or None
    date_from = request.args.get("date_from") or None
    date_to = request.args.get("date_to") or None
    rows = support_db.get_feedback(module=module, status=status, source=source,
                                    date_from=date_from, date_to=date_to)
    return render_template("admin_feedback.html", rows=rows, module=module, status=status,
                            source=source, date_from=date_from, date_to=date_to,
                            new_count=support_db.count_new_feedback())


@support_bp.route("/admin/feedback/<int:feedback_id>/status", methods=["POST"])
@admin_required
def admin_feedback_status(feedback_id):
    status = request.form.get("status")
    if status in ("new", "reviewed", "fixed", "wontfix"):
        support_db.set_feedback_status(feedback_id, status)
    return redirect(url_for("support.admin_feedback"))


@support_bp.route("/admin/errors")
@admin_required
def admin_errors():
    resolved = request.args.get("resolved")
    resolved_bool = {"1": True, "0": False}.get(resolved)
    rows = support_db.get_errors(resolved=resolved_bool)
    return render_template("admin_errors.html", rows=rows, resolved=resolved)


@support_bp.route("/admin/errors/<int:error_id>/resolve", methods=["POST"])
@admin_required
def admin_resolve_error(error_id):
    support_db.resolve_error(error_id)
    return redirect(url_for("support.admin_errors"))


def init_app(app):
    support_db.init_db()
    app.register_blueprint(support_bp)
    app.context_processor(lambda: {
        "support_new_feedback_count": support_db.count_new_feedback() if current_user()
        and current_user()["role"] == "admin" else 0,
    })
