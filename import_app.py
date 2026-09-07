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

from import_parser import parse_invoice_file
from import_combine import combine_by_mark
from import_xls_writer import build_factura_xls
import support.db as support_db
from auth import current_user

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
    errors = []
    for f in files:
        name_lower = f.filename.lower()
        if not name_lower.endswith((".pdf", ".xls", ".xlsx")):
            errors.append(f"{f.filename}: не PDF/XLS-файл, пропущен")
            continue
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{f.filename}")
        f.save(path)
        try:
            data, template_name = parse_invoice_file(path)
            parsed_invoices.append({"filename": f.filename, "data": data, "template": template_name})
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
    _save_pending(token, {"invoices": parsed_invoices})
    session["import_token"] = token

    for e in errors:
        flash(e, "error")

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


def _apply_awb(by_mark, awb_data):
    """Считает AWB-логистику (вес x ставка) для каждой метки из ручного
    ввода (см. factura_awb) и проставляет её в by_mark - общая логика для
    страницы /factura и для выгрузки .xls, чтобы не дублировать формулы."""
    for mark, info in by_mark.items():
        awb = awb_data.get(mark, {})
        pieces = awb.get("pieces")
        gross = awb.get("gross_weight")
        chargeable = awb.get("chargeable_weight") or gross
        rate = awb.get("rate_per_kg")

        weight_per_box = round(gross / pieces, 3) if pieces and gross else None
        total_awb = round(chargeable * rate, 2) if chargeable and rate else None
        awb_per_box = round(total_awb / pieces, 2) if total_awb and pieces else None

        info["awb"] = {
            "pieces": pieces, "gross_weight": gross,
            "chargeable_weight": chargeable, "rate_per_kg": rate,
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
    _apply_awb(by_mark, pending.get("awb", {}))

    return render_template("import_factura.html", by_mark=by_mark)


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
    _apply_awb(by_mark, pending.get("awb", {}))

    buf = io.BytesIO()
    build_factura_xls(buf, mark, by_mark[mark])
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                      download_name=f"Invoice total {mark}.xls",
                      mimetype="application/vnd.ms-excel")
