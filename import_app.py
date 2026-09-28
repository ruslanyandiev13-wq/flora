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

from flask import Blueprint, request, render_template, redirect, url_for, session, flash, send_file

from import_parser import parse_invoice_file, parse_awb_pdf
from import_combine import combine_by_mark
from import_xls_writer import build_factura_xls, build_combined_factura_xls
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
                awb = parse_awb_pdf(path)
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
        flash(f"Авианакладная {awb_doc.get('awb_no') or ''} распознана: "
              f"{awb_doc.get('pieces')} мест, {awb_doc.get('gross_weight')} кг брутто, "
              f"ставка {awb_doc.get('rate_per_kg')} $/кг - вес и ставка подставлены автоматически", "ok")

    return redirect(url_for("import_invoices.review"))


def _combined_boxes(invoices):
    """Сквозная нумерация коробок по всем загруженным инвойсам сразу, как
    в голландском модуле (см. app.py::_combined_boxes)."""
    combined = []
    box_counter = 0
    for inv in invoices:
        for b in inv["data"]["boxes"]:
            box_counter += 1
            combined.append({
                "box_no": box_counter, "box_type": b.get("box_type"),
                "box_size": b.get("box_size"), "items": b["items"],
                "source": inv["filename"], "supplier": inv["data"]["supplier"],
            })
    return combined


@import_bp.route("/review")
def review():
    token = session.get("import_token")
    pending = _load_pending(token)
    if not pending:
        flash("Сессия истекла, загрузите файлы заново", "error")
        return redirect(url_for("import_invoices.index"))

    invoices = pending["invoices"]
    combined = _combined_boxes(invoices)

    # Сверка: сумма, посчитанная парсером по позициям, против суммы,
    # напечатанной в самом инвойсе (totals) - расхождение сигнализирует
    # о возможной ошибке разбора и должно быть видно бухгалтеру сразу.
    for inv in invoices:
        boxes = inv["data"]["boxes"]
        computed_stems = sum(it["stems"] for b in boxes for it in b["items"])
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
    total_fob = sum(inv["computed_fob"] for inv in invoices)

    return render_template(
        "import_review.html",
        invoices=invoices,
        boxes=combined,
        total_stems=total_stems,
        total_fob=total_fob,
    )


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
    """Ручной ввод данных авианакладной (AWB/HAWB) на метку - см. память
    проекта: закупщик подтвердил, что вес отправления берётся ИЗ отдельного
    документа AWB (а не из фермерских инвойсов), а средний вес коробки
    считается как Gross Weight / Pieces. Настоящего PDF самой AWB нет (только
    фото пересланного документа) - разбор текстом не построить без реального
    файла, поэтому эти поля вводятся вручную; когда появится настоящий
    AWB-PDF, здесь можно добавить автоматический парсер по аналогии с
    import_parser.py."""
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
    """Данные по метке из разобранной авианакладной. Брутто/платный вес в AWB
    указан ОДИН на всю накладную, а по меткам есть только число мест - делим
    вес пропорционально числу мест (решение пользователя 2026-09-09; тот же
    принцип усреднения, что использует закупщик)."""
    if not awb_doc:
        return {}
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


def _apply_awb(by_mark, awb_data, awb_doc=None):
    """Считает AWB-логистику (вес x ставка) для каждой метки и проставляет её
    в by_mark - общая логика для страницы /factura и для выгрузки .xls, чтобы
    не дублировать формулы. Источник данных: разобранная авианакладная
    (awb_doc), поверх неё - ручные правки на метку (awb_data), если есть."""
    for mark, info in by_mark.items():
        awb = dict(_awb_from_doc(mark, awb_doc))
        # Транспорт из самих инвойсов (Astoria) имеет приоритет: если он там
        # напечатан, отдельная авианакладная для этой метки не нужна
        # (правка закупщика 2026-09-10). Ставка = стоимость / вес.
        transport = info.get("transport")
        if transport and transport.get("cost_usd"):
            weight = transport.get("weight_kg")
            awb.update({
                "pieces": len(info["boxes"]),
                "gross_weight": weight,
                "chargeable_weight": weight,
                "rate_per_kg": round(transport["cost_usd"] / weight, 4) if weight else None,
                "other_charges": None,
            })
        for key, value in (awb_data.get(mark) or {}).items():
            if value is not None:
                awb[key] = value
        pieces = awb.get("pieces")
        gross = awb.get("gross_weight")
        chargeable = awb.get("chargeable_weight") or gross
        rate = awb.get("rate_per_kg")

        other_charges = awb.get("other_charges") or 0

        weight_per_box = round(gross / pieces, 3) if pieces and gross else None
        # Итог по накладной = фрахт (платный вес x ставка) + прочие сборы
        # перевозчика, т.е. "Total Prepaid" из самой AWB.
        total_awb = (round(chargeable * rate + other_charges, 2)
                      if chargeable and rate else None)
        awb_per_box = round(total_awb / pieces, 2) if total_awb and pieces else None

        info["awb"] = {
            "pieces": pieces, "gross_weight": gross,
            "chargeable_weight": chargeable, "rate_per_kg": rate,
            "other_charges": other_charges or None,
            "weight_per_box": weight_per_box, "total_awb": total_awb,
            "awb_per_box": awb_per_box,
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

    by_mark = combine_by_mark(pending["invoices"])
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

    by_mark = combine_by_mark(pending["invoices"])
    _apply_awb(by_mark, pending.get("awb", {}), pending.get("awb_doc"))

    awb_doc = pending.get("awb_doc") or {}
    name = awb_doc.get("awb_no") or "-".join(by_mark.keys())

    buf = io.BytesIO()
    build_combined_factura_xls(buf, by_mark, awb_doc)
    buf.seek(0)
    _charge_once(token, pending, by_mark, f"Invoice total {name}.xls")
    return send_file(buf, as_attachment=True,
                      download_name=f"Invoice total {name}.xls",
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

    by_mark = combine_by_mark(pending["invoices"])
    if mark not in by_mark:
        flash(f"Метка {mark} не найдена в текущей сессии", "error")
        return redirect(url_for("import_invoices.factura"))
    _apply_awb(by_mark, pending.get("awb", {}), pending.get("awb_doc"))

    buf = io.BytesIO()
    build_factura_xls(buf, mark, by_mark[mark])
    buf.seek(0)
    _charge_once(token, pending, by_mark, f"Invoice total {mark}.xls")
    return send_file(buf, as_attachment=True,
                      download_name=f"Invoice total {mark}.xls",
                      mimetype="application/vnd.ms-excel")
