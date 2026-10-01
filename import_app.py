"""
Flask-роуты для модуля импортных инвойсов (не-Голландия). Параллельно
голландскому app.py - через отдельный Blueprint, чтобы не трогать
голландские роуты/логику (см. ТЗ: "не переписывай голландский модуль,
добавляй новый рядом").

Фаза 1: загрузка PDF/XLS от поставщика -> разбор через import_parser.py ->
ручная проверка результата перед сохранением (аналог review.html в
голландском модуле). Кнопки "Скачать .xls для 1С" здесь пока нет - итоговый
формат выгрузки (Фаза 4) не подтверждён и требует отдельного согласования.
"""
import os
import uuid
import json
import time
import io
import copy
import datetime

from flask import (Blueprint, request, render_template, redirect, url_for, session, flash,
                   send_file, jsonify)

from import_parser import parse_invoice_file, parse_awb_pdf, parse_forwarder_hawb_pdf
from import_combine import combine_by_mark, box_farm_product
import review_edits
import db
import deliveries
from import_xls_writer import build_combined_factura_xls
import support.db as support_db
from auth import current_user
from billing import db as billing_db
from billing.charge import charge_for_invoice

import_bp = Blueprint("import_invoices", __name__, url_prefix="/import")

UPLOAD_DIR = None  # инициализируется в init_app()
PENDING = {}
PENDING_DIR = os.path.join(os.path.dirname(__file__), "data", "pending_import")
PENDING_MAX_AGE_DAYS = 14


def init_app(app, upload_dir):
    global UPLOAD_DIR
    UPLOAD_DIR = upload_dir
    os.makedirs(PENDING_DIR, exist_ok=True)
    app.register_blueprint(import_bp)


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


def _cleanup_old_pending():
    cutoff = time.time() - PENDING_MAX_AGE_DAYS * 86400
    for fn in os.listdir(PENDING_DIR):
        path = os.path.join(PENDING_DIR, fn)
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
        except OSError:
            pass


def _page_count(path):
    """Число страниц PDF - влияет на стоимость обработки в биллинге."""
    if not path.lower().endswith(".pdf"):
        return 1
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return len(pdf.pages)
    except Exception:
        return 1


def _charge_once(token, pending, by_mark, result_name):
    """Списывает токены за партию - ОДИН раз на сессию загрузки, по записи на
    каждую исходную накладную (как в голландском модуле). Повторные скачивания
    того же файла не тарифицируются: отметка о списании хранится в самой
    pending-сессии."""
    if pending.get("charged"):
        return
    user = current_user()
    uploaded_by = user["username"] if user else "?"
    awb_doc = pending.get("awb_doc")

    # Какие накладные попали в объединение с применением правила MIX -
    # это отдельная позиция тарифа.
    mix_sources = set()
    for info in by_mark.values():
        for box in info["boxes"]:
            if any(it.get("merged_from") or it.get("renamed_from") for it in box["items"]):
                mix_sources.add(box.get("source"))

    charged = []
    for i, inv in enumerate(pending["invoices"]):
        line_items_count = sum(len(b["items"]) for b in inv["data"]["boxes"])
        result = charge_for_invoice(billing_db.DEFAULT_ORG_ID, "import", {
            "uploaded_by": uploaded_by,
            "supplier_template": inv["template"],
            "page_count": inv.get("page_count", 1),
            "line_items_count": line_items_count,
            # Авианакладная разбирается автоматически и одна на всю партию -
            # относим её к первой накладной, чтобы не тарифицировать дважды.
            "awb_structured": bool(awb_doc) and i == 0,
            "mix_rule_applied": inv["filename"] in mix_sources,
            "status": "processed",
            "result_file_path": result_name,
            "source_filename": inv["filename"],
            "source_file_path": inv.get("upload_path"),
            "parsed_data": inv["data"],
        })
        charged.append(result["invoice_id"])

    pending["charged"] = charged
    _save_pending(token, pending)


@import_bp.route("/")
def index():
    return render_template("import_index.html")


@import_bp.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("invoices")
    if not files or not files[0].filename:
        flash("Выберите хотя бы один файл (PDF или XLS)", "error")
        return redirect(url_for("import_invoices.index"))

    parsed_invoices = []
    awb_doc = None
    errors = []
    for f in files:
        name_lower = f.filename.lower()
        if not name_lower.endswith((".pdf", ".xls", ".xlsx")):
            errors.append(f"{f.filename}: не PDF/XLS-файл, пропущен")
            continue
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{f.filename}")
        f.save(path)
        try:
            # Авианакладная - один документ на всю партию, а не инвойс фермы:
            # из неё берутся вес/ставка/число мест по меткам (раньше их
            # вводили руками, а сам файл AWB ошибочно разбирался как инвойс
            # и давал пустой результат).
            if name_lower.endswith(".pdf"):
                # HAWB тоже содержит AIR WAYBILL/SHIPPER, но в ней нет
                # тарифа. Общий AWB-парсер узнаёт бланк и возвращает пустые
                # веса/метки, поэтому сначала проверяем специальный формат.
                awb = _awb_doc_from_hawb(parse_forwarder_hawb_pdf(path)) or parse_awb_pdf(path)
                if awb:
                    awb["filename"] = f.filename
                    awb_doc = awb
                    continue
            data, template_name = parse_invoice_file(path)
            parsed_invoices.append({
                "filename": f.filename, "data": data, "template": template_name,
                "page_count": _page_count(path), "upload_path": path,
            })
        except Exception as e:
            errors.append(f"{f.filename}: {e}")
            user = current_user()
            support_db.log_error("import", str(e), uploaded_by=user["username"] if user else None,
                                  source_filename=f.filename, error_type="parse_error")

    if not parsed_invoices:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("import_invoices.index"))

    _cleanup_old_pending()
    token = uuid.uuid4().hex
    _save_pending(token, {"invoices": parsed_invoices, "awb_doc": awb_doc})
    session["import_token"] = token

    for e in errors:
        flash(e, "error")
    if awb_doc:
        flash(f"Авианакладная {awb_doc.get('awb_no') or ''} распознана"
              f"{' по скану (проверьте цифры)' if awb_doc.get('ocr') else ''}: "
              f"{awb_doc.get('pieces')} мест, {awb_doc.get('gross_weight')} кг брутто, "
              f"ставка {awb_doc.get('rate_per_kg')} $/кг - вес и ставка подставлены автоматически", "ok")

    return redirect(url_for("import_invoices.review"))


def _edited_invoices(pending, keep_original=False):
    """Инвойсы сессии с ручными правками со страницы /import/review.

    Правки накладываются на РАЗОБРАННЫЕ позиции, до сборки по меткам -
    поэтому правила MIX/MOON MIX/верхней ростовки дальше работают уже по
    исправленным данным. Исходные pending["invoices"] не меняются.
    keep_original - оставить у исправленных полей исходные значения
    ("original") для подсветки на странице; в factura они не нужны."""
    invoices = copy.deepcopy(pending["invoices"])
    edits = pending.get("edits") or {}
    box_counter = 0
    for inv in invoices:
        for b in inv["data"]["boxes"]:
            box_counter += 1
            review_edits.apply_box_edits(b, edits.get(str(box_counter)), review_edits.IMPORT)
            if not keep_original:
                b.pop("original", None)
                b["items"] = [review_edits.strip_original(it) for it in b["items"]]
    return invoices


def _combined_boxes(invoices):
    """Сквозная нумерация коробок по всем загруженным инвойсам сразу, как
    в голландском модуле (см. app.py::_combined_boxes). FARM/PRODUCT - как
    они попадут в factura."""
    combined = []
    box_counter = 0
    for inv in invoices:
        for b in inv["data"]["boxes"]:
            box_counter += 1
            farm, product = box_farm_product(inv, b)
            combined.append({
                "box_no": box_counter, "box_type": b.get("box_type"),
                "box_size": b.get("box_size"), "items": b["items"],
                "farm": farm, "product": product, "original": b.get("original"),
                "source": inv["filename"], "supplier": inv["data"]["supplier"],
            })
    return combined


def _review_totals(invoices):
    """Сверка: сумма, посчитанная парсером по позициям, против суммы,
    напечатанной в самом инвойсе (totals) - расхождение сигнализирует
    о возможной ошибке разбора и должно быть видно бухгалтеру сразу.
    Дописывает поля сверки в каждый инвойс, возвращает общие стебли и сумму."""
    for inv in invoices:
        boxes = inv["data"]["boxes"]
        computed_stems = sum(it["stems"] or 0 for b in boxes for it in b["items"])
        computed_fob = sum((it["total"] or 0) for b in boxes for it in b["items"])
        inv["computed_stems"] = round(computed_stems, 2)
        inv["computed_fob"] = round(computed_fob, 2)
        printed_stems = inv["data"]["totals"].get("total_stems")
        printed_fob = inv["data"]["totals"].get("total_fob")
        inv["stems_mismatch"] = (printed_stems is not None
                                  and abs(printed_stems - computed_stems) > 0.5)
        inv["fob_mismatch"] = (printed_fob is not None
                                and abs(printed_fob - computed_fob) > 0.5)
    total_stems = sum(inv["computed_stems"] for inv in invoices)
    total_fob = round(sum(inv["computed_fob"] for inv in invoices), 2)
    return total_stems, total_fob


@import_bp.route("/review")
def review():
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    invoices = _edited_invoices(pending, keep_original=True)
    combined = _combined_boxes(invoices)
    total_stems, total_fob = _review_totals(invoices)

    return render_template(
        "import_review.html",
        invoices=invoices,
        boxes=combined,
        total_stems=total_stems,
        total_fob=total_fob,
        edits_count=review_edits.count_edits(pending.get("edits")),
    )


@import_bp.route("/review/edit", methods=["POST"])
def review_edit():
    """Правка одной клетки таблицы позиций (вызывается со страницы
    /import/review). Правки хранятся в pending-сессии и попадают в factura."""
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        return jsonify(error="Сессия истекла, загрузите файлы заново"), 410

    edits = pending.setdefault("edits", {})
    try:
        # Сравниваем с исходными значениями этой клетки: вернули их - правка
        # снимается.
        box_no, item_idx, name = review_edits.record_edit(
            edits, request.get_json(silent=True) or {},
            _combined_boxes(pending["invoices"]), review_edits.IMPORT)
    except ValueError as e:
        return jsonify(error=str(e)), 400
    _save_pending(token, pending)

    invoices = _edited_invoices(pending, keep_original=True)
    combined = _combined_boxes(invoices)
    total_stems, total_fob = _review_totals(invoices)
    response = review_edits.edit_response(combined, box_no, item_idx, name, edits,
                                          review_edits.IMPORT)
    return jsonify(
        **response,
        totals={"total-stems": f"{total_stems:g}", "total-fob": f"{total_fob:.2f}",
                **{f"inv-stems-{i}": f"{inv['computed_stems']:g}" for i, inv in enumerate(invoices)},
                **{f"inv-fob-{i}": f"{inv['computed_fob']:.2f}" for i, inv in enumerate(invoices)}},
        # Правка коробки или сорта может поменять PRODUCT и сверку с
        # инвойсом - такие клетки проще показать после перезагрузки.
        reload=item_idx is None or name == "variety",
    )


@import_bp.route("/review/edits/reset", methods=["POST"])
def review_edits_reset():
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))
    pending.pop("edits", None)
    _save_pending(token, pending)
    flash("Ручные правки сброшены - позиции снова как в инвойсах", "ok")
    return redirect(url_for("import_invoices.review"))


def _to_float_or_none(s):
    s = (s or "").strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _to_int_or_none(s):
    f = _to_float_or_none(s)
    return int(f) if f is not None else None


@import_bp.route("/factura/awb", methods=["POST"])
def factura_awb():
    """Ручные дополнения и исправления автоматически прочитанной перевозки
    на метку. Пустые поля не затирают значения из документов."""
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    mark = request.form.get("mark", "")
    awb = pending.setdefault("awb", {})
    awb[mark] = {
        "pieces": _to_int_or_none(request.form.get("pieces")),
        "gross_weight": _to_float_or_none(request.form.get("gross_weight")),
        "chargeable_weight": _to_float_or_none(request.form.get("chargeable_weight")),
        "rate_per_kg": _to_float_or_none(request.form.get("rate_per_kg")),
    }
    _save_pending(token, pending)
    return redirect(url_for("import_invoices.factura"))


def _awb_from_doc(mark, awb_doc):
    """Данные по метке из разобранной авианакладной.

    Если в AWB есть отдельная накладная на метку (house AWB на следующих
    страницах) - берём её вес, тариф и сборы как есть: так считает закупщик
    (правка 2026-09-28, у DAMIR 415 кг по его накладной, а не 332 кг по доле
    мест). Иначе - как раньше: общий вес делим пропорционально числу мест
    (решение пользователя 2026-09-09)."""
    if not awb_doc:
        return {}
    house = (awb_doc.get("houses") or {}).get(mark)
    if house:
        return {key: house.get(key) for key in
                ("pieces", "gross_weight", "chargeable_weight", "rate_per_kg", "other_charges")}
    mark_info = (awb_doc.get("marks") or {}).get(mark)
    total_pieces = awb_doc.get("pieces")
    if not mark_info or not total_pieces:
        return {}
    share = (mark_info.get("pieces") or 0) / total_pieces
    gross = awb_doc.get("gross_weight")
    chargeable = awb_doc.get("chargeable_weight")
    other = awb_doc.get("other_charges")
    return {
        "pieces": mark_info.get("pieces"),
        "gross_weight": round(gross * share, 2) if gross else None,
        "chargeable_weight": round(chargeable * share, 2) if chargeable else None,
        "rate_per_kg": awb_doc.get("rate_per_kg"),
        # Прочие сборы перевозчика ("Total Other Charges Due Carrier") входят
        # в "Total Prepaid" - итог по накладной должен включать их, а не
        # только фрахт (правка закупщика 2026-09-10). Делим так же, как вес -
        # пропорционально числу мест.
        "other_charges": round(other * share, 2) if other else None,
    }


def _awb_doc_from_hawb(hawb):
    """HAWB форвардера (в т.ч. скан IFC из Боготы, правка закупщика
    2026-10-01) как авианакладная партии. Ставки в HAWB нет ("AS AGREED") -
    берём ставку за кг из справочника поставок по стране вылета."""
    if not hawb or not hawb.get("mark"):
        return None
    country = hawb.get("origin_country") or "ecuador"
    rate = float(db.get_setting(f"delivery_rate_kg_{country}", 8 if country == "colombia" else 8.1))
    return {
        "awb_no": hawb.get("mawb"), "hawb_no": hawb.get("hawb"),
        "pieces": hawb.get("pieces"), "gross_weight": hawb.get("gross_weight"),
        "chargeable_weight": hawb.get("chargeable_weight") or hawb.get("gross_weight"),
        "rate_per_kg": rate, "other_charges": None,
        "marks": {hawb["mark"]: {"pieces": hawb.get("pieces"), "full_boxes": hawb.get("total_full")}},
        "houses": {}, "flight_date": hawb.get("flight_date"), "airport": hawb.get("airport"),
        "ocr": hawb.get("ocr", False),
    }


def _delivery_date(pending):
    """Дата поставки на склад для Invoice total (правка закупщика 2026-10-01):
    московский самолёт - ближайшее воскресенье, через Амстердам (Saftec и т.п.)
    - среда следующей недели после вылета. Дата вылета - из AWB, без AWB -
    самая поздняя дата инвойса."""
    awb_doc = pending.get("awb_doc") or {}
    flight = awb_doc.get("flight_date")
    if not flight:
        dates = [deliveries.parse_date(inv["data"].get("invoice_date"), inv["template"])
                 for inv in pending["invoices"]]
        dates = [d for d in dates if d]
        flight = max(dates) if dates else None
    rule = (deliveries.moscow_delivery_date if awb_doc.get("airport") in ("SVO", "VKO", "DME")
            else deliveries.amsterdam_delivery_date)
    try:
        return rule(datetime.date.fromisoformat(flight))
    except (TypeError, ValueError):
        return None


def _invoice_total_filename(label, delivery_date):
    date_prefix = delivery_date.strftime("%d.%m ") if delivery_date else ""
    return f"Invoice total {date_prefix}{label}.xls"


def _apply_awb(by_mark, awb_data, awb_doc=None):
    """Считает AWB-логистику (вес x ставка) для каждой метки и проставляет её
    в by_mark - общая логика для страницы /factura и для выгрузки .xls, чтобы
    не дублировать формулы. Источник данных: разобранная авианакладная
    (awb_doc), поверх неё - ручные правки на метку (awb_data), если есть."""
    for mark, info in by_mark.items():
        awb = dict(_awb_from_doc(mark, awb_doc))
        # Транспорт из самих инвойсов (Astoria/брокер) имеет приоритет: если он там
        # напечатан, отдельная авианакладная для этой метки не нужна
        # (правка закупщика 2026-09-10). Ставка = стоимость / вес.
        transport = info.get("transport")
        invoice_cost = transport.get("cost_usd") if transport else None
        if transport:
            awb["pieces"] = len(info["boxes"])
            for field in ("gross_weight", "chargeable_weight"):
                weight = transport.get(field)
                if weight is not None:
                    awb[field] = weight
            if invoice_cost is not None:
                weight = awb.get("chargeable_weight") or awb.get("gross_weight")
                awb["rate_per_kg"] = round(invoice_cost / weight, 4) if weight else None
                awb["other_charges"] = None  # TOTAL AWB уже включает сборы.
        # Не пересчитываем напечатанную сумму из округлённой ставки при
        # простом сохранении предзаполненной формы. Изменённые вес/тариф
        # по-прежнему дают закупщику возможность пересчитать перевозку.
        base_weight = awb.get("chargeable_weight") or awb.get("gross_weight")
        base_rate = awb.get("rate_per_kg")
        base_other = awb.get("other_charges") or 0
        overrides = awb_data.get(mark) or {}
        for key, value in overrides.items():
            if value is not None:
                awb[key] = value
        pieces = awb.get("pieces")
        gross = awb.get("gross_weight")
        chargeable = awb.get("chargeable_weight") or gross
        rate = awb.get("rate_per_kg")

        other_charges = awb.get("other_charges") or 0
        cost_changed = ((chargeable, rate, other_charges) != (base_weight, base_rate, base_other))
        if invoice_cost is not None and base_rate is None and overrides.get("rate_per_kg") is None:
            # При известной сумме добавление недостающего веса уточняет
            # производную ставку, а не отменяет напечатанный TOTAL AWB.
            rate = round(invoice_cost / chargeable, 4) if chargeable else None
            cost_changed = other_charges != base_other

        weight_per_box = round(gross / pieces, 3) if pieces and gross else None
        # Итог по накладной = фрахт (платный вес x ставка) + прочие сборы
        # перевозчика, т.е. "Total Prepaid" из самой AWB.
        if invoice_cost is not None and not cost_changed:
            total_awb = invoice_cost
        else:
            total_awb = (round(chargeable * rate + other_charges, 2)
                         if chargeable is not None and rate is not None else None)
        source = ("invoice" if invoice_cost is not None and not cost_changed else
                  "manual" if cost_changed else "awb" if awb_doc else None)
        awb_per_box = round(total_awb / pieces, 2) if total_awb is not None and pieces else None

        info["awb"] = {
            "pieces": pieces, "gross_weight": gross,
            "chargeable_weight": chargeable, "rate_per_kg": rate,
            # "ставка за кг" в файле закупщика - всё, что заплачено за
            # перевозку, на килограмм: итог AWB / платный вес (3.2591 при
            # тарифе 3.25 - сборы AWC тоже в ней).
            "rate_all_in": round(total_awb / chargeable, 4) if total_awb is not None and chargeable else None,
            "other_charges": other_charges or None,
            "weight_per_box": weight_per_box, "total_awb": total_awb,
            "awb_per_box": awb_per_box,
            "source": source,
            "box_count_mismatch": pieces is not None and pieces != len(info["boxes"]),
        }
        if awb_per_box is not None:
            for b in info["boxes"]:
                b["awb_cost"] = awb_per_box


@import_bp.route("/factura")
def factura():
    """Фаза 3: сборка всех загруженных в этой сессии инвойсов в черновик
    по меткам (mark) - см. import_combine.py. Плюс AWB-логистика (Фаза 4,
    вес x ставка) - вес и ставка вводятся вручную на метку (см. factura_awb
    выше), средний вес коробки = Gross Weight / Pieces, стоимость AWB на
    коробку = (Chargeable Weight x rate) / Pieces (равномерный сплит - т.к.
    вес коробки уже усреднён, это эквивалентно пропорциональному сплиту по
    весу, который использует сам закупщик)."""
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    by_mark = combine_by_mark(_edited_invoices(pending))
    _apply_awb(by_mark, pending.get("awb", {}), pending.get("awb_doc"))

    return render_template("import_factura.html", by_mark=by_mark)


@import_bp.route("/factura/download")
def factura_download_all():
    """Единый Invoice total на всю партию (одна авианакладная): все метки в
    одной таблице, метка - отдельной колонкой. Подтверждено закупщиком
    2026-09-09; выгрузка по отдельной метке (ниже) тоже осталась."""
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    by_mark = combine_by_mark(_edited_invoices(pending))
    _apply_awb(by_mark, pending.get("awb", {}), pending.get("awb_doc"))

    awb_doc = pending.get("awb_doc") or {}
    name = awb_doc.get("awb_no") or "-".join(by_mark.keys())
    delivery_date = _delivery_date(pending)
    filename = _invoice_total_filename(name, delivery_date)

    buf = io.BytesIO()
    build_combined_factura_xls(buf, by_mark, awb_doc, delivery_date=delivery_date)
    buf.seek(0)
    _charge_once(token, pending, by_mark, filename)
    return send_file(buf, as_attachment=True,
                      download_name=filename,
                      mimetype="application/vnd.ms-excel")


@import_bp.route("/factura/<mark>/download")
def factura_download(mark):
    """Фаза 4: выгрузка .xls по метке - формат подтверждён реальным примером
    бухгалтера (см. import_xls_writer.py)."""
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    by_mark = combine_by_mark(_edited_invoices(pending))
    if mark not in by_mark:
        flash(f"Метка {mark} не найдена в текущей сессии", "error")
        return redirect(url_for("import_invoices.factura"))
    _apply_awb(by_mark, pending.get("awb", {}), pending.get("awb_doc"))

    # Тот же формат, что у общего файла, только с одной меткой. Итог AWB -
    # доля метки (из _apply_awb), а не вся накладная.
    awb_doc = pending.get("awb_doc") or {}
    delivery_date = _delivery_date(pending)
    filename = _invoice_total_filename(mark, delivery_date)
    buf = io.BytesIO()
    build_combined_factura_xls(buf, {mark: by_mark[mark]}, {"awb_no": awb_doc.get("awb_no")},
                               delivery_date=delivery_date)
    buf.seek(0)
    _charge_once(token, pending, by_mark, filename)
    return send_file(buf, as_attachment=True,
                      download_name=filename,
                      mimetype="application/vnd.ms-excel")
