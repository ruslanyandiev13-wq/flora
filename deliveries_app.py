"""
Раздел «Поставки» (/deliveries) - веб-часть. Логика раскладки коробок по
рейсам - в deliveries.py; здесь загрузка документов, страницы и файлы.

Документы копятся по метке в БД (delivery_docs), а не в одноразовой
сессии, как в «Импорте»: инвойсы и HAWB приходят в разные дни.
"""
import json
import os
import uuid

from flask import Blueprint, request, render_template, redirect, url_for, flash, send_file

import db
import deliveries
from auth import current_user
from import_parser import parse_invoice_file, parse_forwarder_hawb_pdf, parse_awb_pdf
from timeutil import now_str
import support.db as support_db

deliveries_bp = Blueprint("deliveries", __name__, url_prefix="/deliveries")
UPLOAD_DIR = None


def init_app(app, upload_dir):
    global UPLOAD_DIR
    UPLOAD_DIR = upload_dir
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    app.register_blueprint(deliveries_bp)


def _username():
    user = current_user()
    return user["username"] if user else "?"


@deliveries_bp.route("/")
def index():
    return render_template("deliveries_index.html", marks=db.delivery_marks())


@deliveries_bp.route("/upload", methods=["POST"])
def upload():
    files = [f for f in request.files.getlist("docs") if f and f.filename]
    if not files:
        flash("Выберите файлы: инвойсы ферм и HAWB форвардера", "error")
        return redirect(request.referrer or url_for("deliveries.index"))

    saved_marks, counts = [], {"invoice": 0, "hawb": 0}
    for f in files:
        name = f.filename
        if not name.lower().endswith((".pdf", ".xls", ".xlsx")):
            flash(f"{name}: нужен PDF или XLS, файл пропущен", "error")
            continue
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{os.path.basename(name)}")
        f.save(path)
        try:
            hawb = parse_forwarder_hawb_pdf(path) if name.lower().endswith(".pdf") else None
            if hawb:
                if not hawb.get("mark") or not hawb.get("growers"):
                    raise ValueError("в HAWB не нашлись метка или строки ферм")
                db.save_delivery_doc(hawb["mark"], "hawb", hawb.get("hawb") or name, name,
                                     "forwarder_hawb", json.dumps(hawb), now_str(), _username())
                saved_marks.append(hawb["mark"])
                counts["hawb"] += 1
                continue
            if name.lower().endswith(".pdf") and parse_awb_pdf(path):
                raise ValueError("это общая авианакладная (AWB). Для поставок нужны HAWB "
                                 "форвардера с меткой и строками «доля = ферма»")
            data, template = parse_invoice_file(path)
            if not data.get("mark"):
                raise ValueError("в инвойсе не нашлась метка")
            if not data.get("boxes"):
                raise ValueError("в инвойсе не нашлось ни одной коробки")
            doc_key = f"{template}:{data.get('invoice_no') or name}"
            db.save_delivery_doc(data["mark"], "invoice", doc_key, name, template,
                                 json.dumps(data), now_str(), _username())
            saved_marks.append(data["mark"])
            counts["invoice"] += 1
        except Exception as e:
            flash(f"{name}: {e}", "error")
            support_db.log_error("deliveries", str(e), uploaded_by=_username(),
                                  source_filename=name, error_type="parse_error")
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    if counts["invoice"] or counts["hawb"]:
        flash(f"Загружено инвойсов: {counts['invoice']}, HAWB: {counts['hawb']}. "
              "Повторно загруженные документы заменили старые версии.", "ok")
    marks = sorted(set(saved_marks))
    if len(marks) == 1:
        return redirect(url_for("deliveries.mark_view", mark=marks[0]))
    return redirect(url_for("deliveries.index"))


@deliveries_bp.route("/<mark>")
def mark_view(mark):
    model = deliveries.build(mark)
    if not model["docs"]:
        flash(f"По метке {mark} документов нет", "error")
        return redirect(url_for("deliveries.index"))
    return render_template("deliveries_mark.html", m=model, fmt_date=deliveries.fmt_date,
                           fmt_day=deliveries.fmt_day, units=deliveries.UNITS,
                           country_names=deliveries.COUNTRY_NAMES)


@deliveries_bp.route("/<mark>/alias", methods=["POST"])
def alias(mark):
    """«Это та же ферма»: написание из HAWB -> ферма из инвойсов. Сохраняется
    в справочник, раскладка пересчитывается сразу."""
    name = request.form.get("alias", "").strip()
    grower = request.form.get("grower", "").strip()
    if name and grower:
        db.upsert_grower_alias(deliveries.normalize_name(name) or name, grower)
        flash(f"Запомнил: «{name}» — это ферма {grower.upper()}. Коробки разложены заново.", "ok")
    return redirect(url_for("deliveries.mark_view", mark=mark))


@deliveries_bp.route("/<mark>/pin", methods=["POST"])
def pin(mark):
    """Ручная привязка коробки к рейсу ("" - вернуть автоматику)."""
    box_key = request.form.get("box_key", "")
    hawb_id = request.form.get("hawb_id", "")
    db.set_delivery_pin(box_key, int(hawb_id) if hawb_id.isdigit() else None)
    flash("Привязка коробки сохранена" if hawb_id else "Коробка снова раскладывается автоматически", "ok")
    return redirect(url_for("deliveries.mark_view", mark=mark) + "#" + request.form.get("anchor", ""))


@deliveries_bp.route("/doc/<int:doc_id>/delete", methods=["POST"])
def delete_doc(doc_id):
    mark = request.form.get("mark", "")
    db.delete_delivery_doc(doc_id)
    flash("Документ удалён", "ok")
    if mark and db.get_delivery_docs(mark):
        return redirect(url_for("deliveries.mark_view", mark=mark))
    return redirect(url_for("deliveries.index"))


def _find_delivery(mark, delivery_id):
    model = deliveries.build(mark)
    return next((d for d in model["deliveries"] if d["id"] == delivery_id), None)


@deliveries_bp.route("/<mark>/<delivery_id>/acceptance.xls")
def acceptance(mark, delivery_id):
    delivery = _find_delivery(mark, delivery_id)
    if not delivery:
        flash("Поставка не найдена - документы могли измениться", "error")
        return redirect(url_for("deliveries.mark_view", mark=mark))
    return send_file(deliveries.acceptance_xls(mark, delivery), as_attachment=True,
                     download_name=f"Приёмка {mark} {deliveries.fmt_date(delivery['date'])} "
                                   f"{delivery['airport'] or ''}.xls",
                     mimetype="application/vnd.ms-excel")


@deliveries_bp.route("/<mark>/<delivery_id>/invoice-total.xls")
def invoice_total(mark, delivery_id):
    delivery = _find_delivery(mark, delivery_id)
    if not delivery or not delivery["boxes_found"]:
        flash("В поставке нет коробок из инвойсов - Invoice total не из чего собрать", "error")
        return redirect(url_for("deliveries.mark_view", mark=mark))
    return send_file(deliveries.invoice_total_xls(delivery), as_attachment=True,
                     download_name=f"Invoice total {mark} {deliveries.fmt_date(delivery['date'])} "
                                   f"{delivery['airport'] or ''}.xls",
                     mimetype="application/vnd.ms-excel")
