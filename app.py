import os
import uuid
import json
import time
import datetime
import tempfile

import pdfplumber
from flask import (Flask, request, render_template, redirect, url_for,
                    send_file, session, flash, jsonify)

import db
import assortment
import assortment_import
import review_edits
from pdf_parser import parse_invoice_pdf
from xls_writer import build_xls
from transport import calculate_transport
from currency import fetch_ligovka_rates
import import_app
import deliveries_app
import deliveries
import auth
from billing import db as billing_db
from billing import routes as billing_routes
from billing.charge import charge_for_invoice
import support.db as support_db
import support.routes as support_routes

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

UPLOAD_DIR = tempfile.mkdtemp(prefix="flora_uploads_")
GENERATED_DIR = tempfile.mkdtemp(prefix="flora_generated_")

# Разобранные поставки по session-токену. Держим и в памяти процесса (быстрый
# доступ), и дублируем на диск в data/pending/ - иначе переход на другую
# страницу и обратно (или рестарт сервиса при деплое) сбрасывает загруженные
# инвойсы и всё приходится грузить заново.
PENDING = {}
PENDING_DIR = os.path.join(os.path.dirname(__file__), "data", "pending")
PENDING_MAX_AGE_DAYS = 14


def _pending_path(token):
    return os.path.join(PENDING_DIR, f"{token}.json")


def _save_pending(token, data):
    PENDING[token] = data
    with open(_pending_path(token), "w", encoding="utf-8") as f:
        json.dump(data, f)


def _load_pending(token):
    if not token:
        return None
    if token in PENDING:
        return PENDING[token]
    path = _pending_path(token)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        PENDING[token] = data
        return data
    return None


def _uploaded_by():
    user = auth.current_user()
    return user["username"] if user else "?"


def _cleanup_old_pending():
    cutoff = time.time() - PENDING_MAX_AGE_DAYS * 86400
    for fn in os.listdir(PENDING_DIR):
        path = os.path.join(PENDING_DIR, fn)
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
        except OSError:
            pass


os.makedirs(PENDING_DIR, exist_ok=True)
db.init_db()

IMPORT_UPLOAD_DIR = tempfile.mkdtemp(prefix="flora_import_uploads_")
import_app.init_app(app, IMPORT_UPLOAD_DIR)
deliveries_app.init_app(app, tempfile.mkdtemp(prefix="flora_delivery_uploads_"))

auth.init_app(app)
billing_routes.init_app(app)
support_routes.init_app(app)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    balance = billing_db.get_balance()
    if balance and balance["current_balance"] <= 0:
        flash("Баланс токенов исчерпан - новая обработка недоступна. "
              "Обратитесь к администратору для пополнения.", "error")
        return redirect(url_for("index"))

    files = request.files.getlist("invoices")
    if not files or not files[0].filename:
        flash("Выберите хотя бы один PDF-файл", "error")
        return redirect(url_for("index"))

    parsed_invoices = []
    errors = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            errors.append(f"{f.filename}: не PDF-файл, пропущен")
            continue
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{f.filename}")
        f.save(path)
        try:
            data, template_name = parse_invoice_pdf(path)
            with pdfplumber.open(path) as pdf:
                page_count = len(pdf.pages)
            parsed_invoices.append({"filename": f.filename, "data": data, "template": template_name,
                                     "page_count": page_count, "upload_path": path})
        except Exception as e:
            errors.append(f"{f.filename}: {e}")
            support_db.log_error("netherlands", str(e), uploaded_by=_uploaded_by(),
                                  source_filename=f.filename, error_type="parse_error")

    if not parsed_invoices:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("index"))

    _cleanup_old_pending()
    token = uuid.uuid4().hex
    _save_pending(token, {"invoices": parsed_invoices})
    session["token"] = token

    for e in errors:
        flash(e, "error")

    return redirect(url_for("review"))


def _combined_boxes(invoices):
    """Объединяет коробки всех загруженных инвойсов в один пронумерованный список.

    Здесь же названия сортов приводятся к справочнику ассортимента бухгалтера
    (см. assortment.py): в паклисте MH Flowers к сорту иногда дописана
    плантация, и решить, часть это названия или мусор, можно только по нему."""
    lookup = assortment.load_lookup()
    combined = []
    box_counter = 0
    for inv in invoices:
        for b in inv["data"]["boxes"]:
            box_counter += 1
            items = [{
                "aantal": it["aantal"],
                "omschrijving": assortment.canonical_name(
                    it.get("omschrijving_printed"), it["omschrijving"], lookup),
                "prijs": it["prijs"], "lengte": it["lengte"], "gew": it["gew"],
                "bedrag": it["bedrag"],
            } for it in b["items"]]
            combined.append({"box_no": box_counter, "fust": b["fust"], "items": items,
                              "source": inv["filename"]})
    return combined


def _reviewed_boxes(pending):
    """Объединённые коробки с ручными правками со страницы /review и сумма
    товара с их учётом (Subtotaal инвойсов + разница по исправленным суммам)."""
    invoices = pending["invoices"]
    combined, money_delta = review_edits.apply_edits(_combined_boxes(invoices),
                                                     pending.get("edits"))
    total_money = sum(inv["data"]["totals"]["subtotaal"] or 0 for inv in invoices)
    return combined, round(total_money + money_delta, 2)


@app.route("/review")
def review():
    token = session.get("token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("index"))

    invoices = pending["invoices"]
    combined, total_money = _reviewed_boxes(pending)
    invoices_money = sum(inv["data"]["totals"]["subtotaal"] or 0 for inv in invoices)
    currencies = {inv["data"]["totals"]["currency"] for inv in invoices}

    tara_mapping = db.get_tara_mapping()
    pallet_cost = float(db.get_setting("pallet_cost_usd", 1650))
    transport = calculate_transport(combined, tara_mapping, pallet_cost)

    recipients = db.get_recipients()

    # Курс USD подтягиваем сразу при открытии страницы (чтобы поле в форме
    # генерации файла уже было предзаполнено) - если ligovka.ru недоступен
    # или сменил вёрстку, не падаем, а даём бухгалтеру вписать курс вручную.
    try:
        usd_rate = fetch_ligovka_rates()["usd_sell_10000_with_markup"]
        currency_error = None
    except Exception as e:
        usd_rate = None
        currency_error = str(e)

    return render_template(
        "review.html",
        invoices=invoices,
        boxes=combined,
        total_money=total_money,
        invoices_money=invoices_money,
        edits_count=review_edits.count_edits(pending.get("edits")),
        currencies=currencies,
        transport=transport,
        pallet_cost=pallet_cost,
        usd_rate=usd_rate,
        currency_error=currency_error,
        recipients=recipients,
        today=datetime.date.today().isoformat(),
        # Неделя считается по дате поставки (правка бухгалтера 2026-09-15:
        # в выгрузке стояла единица от старого шаблона). В форме поле
        # пересчитывается при смене даты, но остаётся редактируемым.
        week_no=datetime.date.today().isocalendar()[1],
    )


@app.route("/review/edit", methods=["POST"])
def review_edit():
    """Правка одной клетки таблицы позиций (вызывается со страницы /review)."""
    token = session.get("token")
    pending = _load_pending(token)
    if not pending:
        return jsonify(error="Сессия истекла, загрузите файлы заново"), 410

    edits = pending.setdefault("edits", {})
    try:
        box_no, item_idx, name = review_edits.record_edit(
            edits, request.get_json(silent=True) or {},
            _combined_boxes(pending["invoices"]))
    except ValueError as e:
        return jsonify(error=str(e)), 400
    _save_pending(token, pending)

    combined, total_money = _reviewed_boxes(pending)
    return jsonify(
        **review_edits.edit_response(combined, box_no, item_idx, name, edits),
        totals={"total-money": f"{total_money:.2f}"},
        # Тара влияет на расчёт логистики - страницу проще перезагрузить.
        reload=item_idx is None,
    )


@app.route("/review/edits/reset", methods=["POST"])
def review_edits_reset():
    token = session.get("token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("index"))
    pending.pop("edits", None)
    _save_pending(token, pending)
    flash("Ручные правки сброшены - позиции снова как в инвойсах", "ok")
    return redirect(url_for("review"))


@app.route("/currency")
def currency_view():
    try:
        rates = fetch_ligovka_rates()
        error = None
    except Exception as e:
        rates = None
        error = str(e)
    return render_template("currency_partial.html", rates=rates, error=error)


@app.route("/generate", methods=["POST"])
def generate():
    token = session.get("token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("index"))

    invoices = pending["invoices"]
    combined, total_money = _reviewed_boxes(pending)

    blad_no = int(request.form.get("blad_no") or 1)
    date_str = request.form.get("date") or datetime.date.today().isoformat()
    date_val = datetime.date.fromisoformat(date_str)
    # Если неделю не заполнили - считаем её по дате поставки (ISO-неделя), а не
    # пишем единицу от старого шаблона: бухгалтер поймала это на файле 13.09.
    week_no = int(request.form.get("week_no") or date_val.isocalendar()[1])
    debnr = request.form.get("debnr", "")
    naam = request.form.get("naam", "")

    pallet_cost_raw = (request.form.get("pallet_cost_usd") or "").strip()
    if pallet_cost_raw:
        pallet_cost = float(pallet_cost_raw.replace(",", "."))
    else:
        pallet_cost = float(db.get_setting("pallet_cost_usd", 1650))

    usd_rate_raw = (request.form.get("usd_rate") or "").strip()
    if usd_rate_raw:
        usd_rate = float(usd_rate_raw.replace(",", "."))
    else:
        try:
            usd_rate = fetch_ligovka_rates()["usd_sell_10000_with_markup"]
        except Exception:
            usd_rate = None

    header = {
        "petia": "PETIA",
        "paklijst_title": "Paklijst Export Unie Flora",
        "week_no": week_no,
        "date": date_val,
        "blad_no": blad_no,
        "debnr": debnr,
        "naam": naam,
    }

    # "original" (исходные значения исправленных полей) в xls не нужен.
    out_boxes = [{"box_no": b["box_no"], "fust": b["fust"],
                  "items": [review_edits.strip_original(it) for it in b["items"]]}
                 for b in combined]
    total_stems = sum(it["aantal"] or 0 for b in combined for it in b["items"])
    out_fname = f"{uuid.uuid4().hex}.xls"
    out_path = os.path.join(GENERATED_DIR, out_fname)
    try:
        build_xls(out_path, out_boxes, header, total_money, total_stems=total_stems,
                  pallet_cost_usd=pallet_cost, usd_rate=usd_rate)
    except Exception as e:
        support_db.log_error("netherlands", str(e), uploaded_by=_uploaded_by(),
                              error_type="exception")
        flash(f"Не удалось сгенерировать файл: {e}", "error")
        return redirect(url_for("review"))

    # generation_context - всё, что нужно, чтобы заново вызвать build_xls и
    # получить тот же файл (нужно для чат-правок, см. support/regenerate.py).
    # Общий для всех накладных этой генерации - при "лёгкой"/"тяжёлой" правке
    # любой из них пересобирается ВЕСЬ объединённый файл.
    generation_context = {
        "boxes": out_boxes, "header": {**header, "date": date_val.isoformat()},
        "total_money": total_money, "pallet_cost_usd": pallet_cost, "usd_rate": usd_rate,
        # Ручные правки со страницы /review - кто и что поменял перед выгрузкой.
        "manual_edits": pending.get("edits") or {}, "edited_by": _uploaded_by(),
    }

    # Токены списываются только после успешной генерации файла (см. ТЗ по
    # биллингу), по одной записи на каждую исходную накладную - даже если
    # несколько PDF объединились в один выгруженный .xls.
    uploaded_by = _uploaded_by()
    document_ids = []
    for inv in invoices:
        line_items_count = sum(len(b["items"]) for b in inv["data"]["boxes"])
        result = charge_for_invoice(billing_db.DEFAULT_ORG_ID, "netherlands", {
            "uploaded_by": uploaded_by,
            "supplier_template": inv["template"],
            "page_count": inv.get("page_count", 1),
            "line_items_count": line_items_count,
            "status": "processed",
            "result_file_path": out_path,
            "source_filename": inv["filename"],
            "source_file_path": inv.get("upload_path"),
            "parsed_data": inv["data"],
            "generation_context": generation_context,
        })
        document_ids.append(result["invoice_id"])

    download_name = f"Голландия для загрузки {date_val.strftime('%d.%m.%Y')}.xls"
    return render_template("result.html", module="netherlands", download_name=download_name,
                            out_fname=out_fname, document_ids=document_ids)


@app.route("/download/<fname>")
def download(fname):
    path = os.path.join(GENERATED_DIR, fname)
    if not os.path.exists(path):
        flash("Файл больше не доступен - сгенерируйте заново", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True,
                      download_name=request.args.get("name") or fname)


# --------------------------------------------------------------------------
# Справочники
# --------------------------------------------------------------------------

@app.route("/dictionaries", methods=["GET", "POST"])
def dictionaries():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add_tara":
            code = request.form.get("fust_code", "").strip()
            label = request.form.get("label", "").strip()
            capacity = request.form.get("capacity_per_pallet", "").strip()
            if code and capacity:
                db.upsert_tara(code, label, float(capacity.replace(",", ".")))
                flash(f"Код тары «{code}» сохранён", "ok")
        elif action == "delete_tara":
            code = request.form.get("fust_code")
            db.delete_tara(code)
            flash(f"Код тары «{code}» удалён", "ok")
        elif action == "set_pallet_cost":
            cost = request.form.get("pallet_cost_usd", "").strip()
            if cost:
                db.set_setting("pallet_cost_usd", float(cost.replace(",", ".")))
                flash("Стоимость палеты сохранена", "ok")
        elif action == "add_recipient":
            label = request.form.get("r_label", "").strip()
            debnr = request.form.get("r_debnr", "").strip()
            naam = request.form.get("r_naam", "").strip()
            if label and debnr and naam:
                db.add_recipient(label, debnr, naam)
                flash(f"Получатель «{label}» сохранён", "ok")
        elif action == "delete_recipient":
            db.delete_recipient(int(request.form.get("recipient_id")))
            flash("Получатель удалён", "ok")
        elif action == "upload_assortment":
            file = request.files.get("assortment_file")
            if not file or not file.filename:
                flash("Выберите файл ассортимента (.xlsx)", "error")
            else:
                try:
                    count = db.replace_assortment(assortment_import.read_xlsx(file))
                    flash(f"Ассортимент загружен: {count} позиций", "ok")
                except Exception as e:
                    flash(f"Не удалось прочитать файл ассортимента: {e}", "error")
        elif action == "add_alias":
            printed = request.form.get("alias_printed", "").strip()
            canonical = request.form.get("alias_canonical", "").strip()
            if printed and canonical:
                db.upsert_variety_alias(printed, canonical)
                flash(f"Замена «{printed}» → «{canonical}» сохранена", "ok")
        elif action == "delete_alias":
            printed = request.form.get("alias_printed")
            db.delete_variety_alias(printed)
            flash(f"Замена «{printed}» удалена", "ok")
        elif action == "add_grower_alias":
            alias = request.form.get("alias", "").strip()
            grower = request.form.get("grower", "").strip()
            if alias and grower:
                db.upsert_grower_alias(alias, grower)
                flash(f"«{alias}» → ферма {grower.upper()}", "ok")
        elif action == "delete_grower_alias":
            alias = request.form.get("alias")
            db.delete_grower_alias(alias)
            flash(f"Написание «{alias}» удалено", "ok")
        elif action == "set_delivery_rates":
            try:
                for country in ("ecuador", "colombia"):
                    value = float(request.form.get(f"rate_{country}", "").replace(",", "."))
                    db.set_setting(f"delivery_rate_kg_{country}", value)
                db.set_setting("delivery_transit_days", int(request.form.get("transit_days") or 1))
                flash("Ставки перевозки сохранены", "ok")
            except ValueError:
                flash("Ставка должна быть числом, например 8.1", "error")
        elif action == "add_suffix":
            suffix = request.form.get("suffix", "").strip()
            if suffix:
                db.add_grower_suffix(suffix)
                flash(f"Приписка «{suffix}» добавлена", "ok")
        elif action == "delete_suffix":
            suffix = request.form.get("suffix")
            db.delete_grower_suffix(suffix)
            flash(f"Приписка «{suffix}» удалена", "ok")
        return redirect(url_for("dictionaries"))

    return render_template(
        "dictionaries.html",
        tara_mapping=db.get_tara_mapping(),
        pallet_cost=db.get_setting("pallet_cost_usd", 1650),
        recipients=db.get_recipients(),
        assortment_count=db.assortment_count(),
        variety_aliases=db.get_variety_aliases(),
        grower_suffixes=db.get_grower_suffixes(),
        grower_aliases=db.get_grower_aliases(),
        rate_ecuador=db.get_setting("delivery_rate_kg_ecuador", "8.1"),
        rate_colombia=db.get_setting("delivery_rate_kg_colombia", "8"),
        transit_days=db.get_setting("delivery_transit_days", "1"),
        known_growers=deliveries.known_growers(),
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
