"""
Роуты раздела "Биллинг / История" (ТЗ, разделы 4-6). Третий, независимый
верхнеуровневый раздел приложения - использует billing/db.py, ничего не
трогает в голландском/импортном модулях, кроме одного вызова
billing.charge_for_invoice() из их роутов generate() (см. ТЗ, раздел 4).
"""
import datetime

from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify

from auth import login_required, admin_required, create_user, list_users, delete_user
from billing import db as billing_db
import timeutil
from billing.pricing import MIN_TOPUP_TOKENS, PRICE_PER_TOKEN_RUB, MONTHLY_TOKEN_PACKAGE

billing_bp = Blueprint("billing", __name__, url_prefix="/billing")


def _balance_context(org_id=billing_db.DEFAULT_ORG_ID):
    balance = billing_db.get_balance(org_id)
    if not balance:
        return {"balance": None, "days_left": None, "low_balance": False}

    since = (timeutil.today() - datetime.timedelta(days=14)).isoformat()
    daily = billing_db.get_daily_spend(org_id, since_date=since)
    total_spent = sum(d["spent"] for d in daily)
    avg_daily = total_spent / 14 if total_spent else 0
    days_left = round(balance["current_balance"] / avg_daily) if avg_daily > 0 else None

    low_balance = balance["current_balance"] < 0.10 * balance["monthly_token_grant"]
    return {"balance": balance, "days_left": days_left, "low_balance": low_balance}


@billing_bp.route("/")
@login_required
def index():
    ctx = _balance_context()
    ctx["user_count"] = len(list_users())
    ctx["storage_used_mb"] = billing_db.get_storage_used_mb()
    return render_template("billing_index.html", **ctx)


@billing_bp.route("/topups")
@login_required
def topups():
    return render_template("billing_topups.html", rows=billing_db.get_topup_history())


@billing_bp.route("/history")
@login_required
def history():
    module = request.args.get("module") or None
    status = request.args.get("status") or None
    uploaded_by = request.args.get("uploaded_by") or None
    complexity_label = request.args.get("complexity_label") or None
    date_from = request.args.get("date_from") or None
    date_to = request.args.get("date_to") or None
    rows = billing_db.get_history(module=module, status=status, uploaded_by=uploaded_by,
                                   complexity_label=complexity_label,
                                   date_from=date_from, date_to=date_to)
    uploaders = billing_db.get_distinct_uploaders()
    return render_template("billing_history.html", rows=rows, module=module, status=status,
                            uploaded_by=uploaded_by, complexity_label=complexity_label,
                            date_from=date_from, date_to=date_to, uploaders=uploaders)


@billing_bp.route("/usage_chart_data")
@login_required
def usage_chart_data():
    balance = billing_db.get_balance()
    period_start = None
    if balance:
        renewal = datetime.date.fromisoformat(balance["plan_renewal_date"])
        period_start = (renewal - datetime.timedelta(days=30)).isoformat()
    daily = billing_db.get_daily_spend(since_date=period_start)

    cumulative = []
    running = 0
    for d in daily:
        running += d["spent"]
        cumulative.append({"day": d["day"], "spent": d["spent"], "cumulative": running})
    return jsonify(cumulative)


@billing_bp.route("/admin/topup", methods=["GET", "POST"])
@admin_required
def admin_topup():
    if request.method == "POST":
        try:
            tokens = int(request.form.get("tokens", "0"))
        except ValueError:
            tokens = 0
        comment = request.form.get("comment", "").strip()
        if tokens < MIN_TOPUP_TOKENS:
            flash(f"Минимальный объём докупки - {MIN_TOPUP_TOKENS} токенов", "error")
        elif not comment:
            flash("Укажите причину/комментарий (например, номер счёта)", "error")
        else:
            billing_db.manual_topup(tokens, comment)
            flash(f"Баланс пополнен на {tokens} токенов", "ok")
        return redirect(url_for("billing.admin_topup"))

    ctx = _balance_context()
    ledger = billing_db.get_ledger(limit=50)
    return render_template("billing_admin_topup.html", ledger=ledger,
                            min_topup=MIN_TOPUP_TOKENS, price_per_token=PRICE_PER_TOKEN_RUB,
                            monthly_package=MONTHLY_TOKEN_PACKAGE, **ctx)


@billing_bp.route("/admin/users", methods=["GET", "POST"])
@admin_required
def admin_users():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "create":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            role = request.form.get("role")
            if not username or len(password) < 8 or role not in ("admin", "user"):
                flash("Логин обязателен, пароль - минимум 8 символов, роль admin/user", "error")
            else:
                try:
                    create_user(username, password, role)
                    flash(f"Пользователь «{username}» создан", "ok")
                except Exception as e:
                    flash(f"Не удалось создать пользователя: {e}", "error")
        elif action == "delete":
            delete_user(int(request.form.get("user_id")))
            flash("Пользователь удалён", "ok")
        return redirect(url_for("billing.admin_users"))

    return render_template("billing_admin_users.html", users=list_users())


def init_app(app):
    billing_db.init_db()
    app.register_blueprint(billing_bp)
    app.context_processor(lambda: {"billing_widget": _balance_context()})
    # Показ отметок времени: "2026-09-11T14:35:02" -> "11.09.2026 14:35" (GMT+3).
    app.jinja_env.filters["dt"] = timeutil.format_dt
    app.jinja_env.globals["TZ_LABEL"] = timeutil.TZ_LABEL
