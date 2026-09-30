"""
Раздел «Поставки» (/deliveries) - веб-часть. Логика раскладки коробок по
рейсам - в deliveries.py; здесь загрузка документов, страницы и файлы.

Документы копятся по метке в БД (delivery_docs), а не в одноразовой
сессии, как в «Импорте»: инвойсы и HAWB приходят в разные дни.
"""
import datetime
import json
import os
import uuid

from flask import Blueprint, request, render_template, redirect, url_for, flash, send_file

import db
import deliveries
from billing import db as billing_db
from billing.charge import charge_for_invoice
from billing.cost_calc import COST_DELIVERY_HAWB
from import_app import _page_count
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


# Документ без явного выбора поставки добавляется к последней поставке своей
# метки, если в неё что-то загружали за эти дни: инвойсы и HAWB одной
# поставки приходят в течение нескольких дней.
AUTO_BATCH_DAYS = 7


def _batches_view():
    batches = db.get_delivery_batches()
    for b in batches:
        b["label"] = deliveries.batch_label(b, batches)
    return batches


@deliveries_bp.route("/")
def index():
    return render_template("deliveries_index.html", batches=_batches_view())


class _BatchChooser:
    """Куда класть документы этой загрузки: выбранная поставка, новая или
    «автоматически». Поставка создаётся только когда в неё реально лёг
    документ - пустых поставок не бывает."""

    def __init__(self, target):
        self.target = target or "auto"
        self.by_mark = {}
        self.notes = []

    def batch_for(self, mark):
        if mark in self.by_mark:
            return self.by_mark[mark]
        batch_id = None
        if self.target.isdigit():
            chosen = db.get_delivery_batch(int(self.target))
            if chosen and chosen["mark"] == mark:
                batch_id = chosen["id"]
            else:
                self.notes.append(f"документы метки {mark} не подходят к выбранной поставке "
                                  f"({chosen['mark'] if chosen else '—'}) - положены отдельно")
        if batch_id is None and self.target != "new":
            cutoff = (datetime.datetime.now() - datetime.timedelta(days=AUTO_BATCH_DAYS)).isoformat()
            recent = [b for b in db.get_delivery_batches()
                      if b["mark"] == mark and (b["last_upload"] or "") >= cutoff]
            if recent:
                batch_id = recent[0]["id"]
        if batch_id is None:
            batch_id = db.create_delivery_batch(mark, now_str())
        self.by_mark[mark] = batch_id
        return batch_id


@deliveries_bp.route("/upload", methods=["POST"])
def upload():
    files = [f for f in request.files.getlist("docs") if f and f.filename]
    if not files:
        flash("Выберите файлы: инвойсы ферм и HAWB форвардера", "error")
        return redirect(request.referrer or url_for("deliveries.index"))

    balance = billing_db.get_balance()
    if balance and balance["current_balance"] <= 0:
        flash("Баланс токенов исчерпан - загрузка документов недоступна. "
              "Обратитесь к администратору для пополнения.", "error")
        return redirect(request.referrer or url_for("deliveries.index"))

    chooser = _BatchChooser(request.form.get("batch"))
    touched, counts, tokens, replaced = set(), {"invoice": 0, "hawb": 0}, 0, 0

    def save(mark, kind, doc_key, name, template, data):
        # Уже загруженный документ остаётся в своей поставке - новую под него не заводим.
        batch_id = db.delivery_doc_batch(mark, kind, doc_key) or chooser.batch_for(mark)
        doc_id, is_new, batch_id = db.save_delivery_doc(
            batch_id, mark, kind, doc_key, name, template,
            json.dumps(data), now_str(), _username())
        touched.add(batch_id)
        counts[kind] += 1
        return doc_id, is_new

    def charge(doc_id, is_new, name, template, metadata):
        """Токены - один раз за документ: замена версии того же инвойса/HAWB
        бесплатна (так же, как повторное скачивание в «Импорте»)."""
        nonlocal tokens, replaced
        if not is_new:
            replaced += 1
            return
        result = charge_for_invoice(billing_db.DEFAULT_ORG_ID, "deliveries", {
            "uploaded_by": _username(), "supplier_template": template,
            "status": "processed", "source_filename": name, **metadata})
        db.set_delivery_doc_billing(doc_id, result["invoice_id"])
        tokens += result["cost"]

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
                doc_id, is_new = save(hawb["mark"], "hawb", hawb.get("hawb") or name, name,
                                      "forwarder_hawb", hawb)
                charge(doc_id, is_new, name, "forwarder_hawb", {
                    "fixed_cost": COST_DELIVERY_HAWB, "page_count": _page_count(path),
                    "line_items_count": len(hawb["growers"]), "awb_free_text": True})
                continue
            if name.lower().endswith(".pdf") and parse_awb_pdf(path):
                raise ValueError("это общая авианакладная (AWB). Для поставок нужны HAWB "
                                 "форвардера с меткой и строками «доля = ферма»")
            data, template = parse_invoice_file(path)
            if not data.get("mark"):
                raise ValueError("в инвойсе не нашлась метка")
            if not data.get("boxes"):
                raise ValueError("в инвойсе не нашлось ни одной коробки")
            doc_id, is_new = save(data["mark"], "invoice", f"{template}:{data.get('invoice_no') or name}",
                                  name, template, data)
            charge(doc_id, is_new, name, template, {
                "page_count": _page_count(path),
                "line_items_count": sum(len(b["items"]) for b in data["boxes"])})
        except Exception as e:
            flash(f"{name}: {e}", "error")
            support_db.log_error("deliveries", str(e), uploaded_by=_username(),
                                  source_filename=name, error_type="parse_error")
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    for note in chooser.notes:
        flash(note[:1].upper() + note[1:], "error")
    if counts["invoice"] or counts["hawb"]:
        batches = _batches_view()
        labels = ", ".join(b["label"] for b in batches if b["id"] in touched)
        flash(f"Загружено инвойсов: {counts['invoice']}, HAWB: {counts['hawb']} → {labels}. "
              f"Списано токенов: {tokens}."
              + (f" Заменено старых версий без списания: {replaced}." if replaced else ""), "ok")
    if len(touched) == 1:
        return redirect(url_for("deliveries.batch_view", batch_id=touched.pop()))
    return redirect(url_for("deliveries.index"))


@deliveries_bp.route("/<mark>")
def mark_view(mark):
    """Старые ссылки вида /deliveries/BESST - на последнюю поставку метки."""
    batch = next((b for b in db.get_delivery_batches() if b["mark"] == mark), None)
    if not batch:
        flash(f"По метке {mark} поставок нет", "error")
        return redirect(url_for("deliveries.index"))
    return redirect(url_for("deliveries.batch_view", batch_id=batch["id"]))


def _model_or_redirect(batch_id):
    model = deliveries.build(batch_id)
    if not model["batch"] or not model["docs"]:
        flash("Поставка не найдена - возможно, все её документы удалены", "error")
        return None
    return model


@deliveries_bp.route("/b/<int:batch_id>")
def batch_view(batch_id):
    model = _model_or_redirect(batch_id)
    if model is None:
        return redirect(url_for("deliveries.index"))
    return render_template("deliveries_mark.html", m=model, fmt_date=deliveries.fmt_date,
                           fmt_day=deliveries.fmt_day, units=deliveries.UNITS,
                           country_names=deliveries.COUNTRY_NAMES)


@deliveries_bp.route("/b/<int:batch_id>/alias", methods=["POST"])
def alias(batch_id):
    """«Это та же ферма»: написание из HAWB -> ферма из инвойсов. Сохраняется
    в справочник, раскладка пересчитывается сразу."""
    name = request.form.get("alias", "").strip()
    grower = request.form.get("grower", "").strip()
    if name and grower:
        db.upsert_grower_alias(deliveries.normalize_name(name) or name, grower)
        flash(f"Запомнил: «{name}» — это ферма {grower.upper()}. Коробки разложены заново.", "ok")
    return redirect(url_for("deliveries.batch_view", batch_id=batch_id))


@deliveries_bp.route("/b/<int:batch_id>/pin", methods=["POST"])
def pin(batch_id):
    """Ручная привязка коробки к рейсу ("" - вернуть автоматику)."""
    box_key = request.form.get("box_key", "")
    hawb_id = request.form.get("hawb_id", "")
    db.set_delivery_pin(box_key, int(hawb_id) if hawb_id.isdigit() else None)
    flash("Привязка коробки сохранена" if hawb_id else "Коробка снова раскладывается автоматически", "ok")
    return redirect(url_for("deliveries.batch_view", batch_id=batch_id) + "#" + request.form.get("anchor", ""))


@deliveries_bp.route("/doc/<int:doc_id>/delete", methods=["POST"])
def delete_doc(doc_id):
    batch_id = request.form.get("batch_id", "")
    db.delete_delivery_doc(doc_id)
    flash("Документ удалён", "ok")
    if batch_id.isdigit() and db.get_delivery_batch(int(batch_id)):
        return redirect(url_for("deliveries.batch_view", batch_id=int(batch_id)))
    return redirect(url_for("deliveries.index"))


def _file_name(prefix, model, delivery):
    label = model["label"].replace(" · ", " ")
    return (f"{prefix} {label} - доставка {deliveries.fmt_date(delivery.get('arrival'))} "
            f"{delivery['airport'] or ''}.xls")


@deliveries_bp.route("/b/<int:batch_id>/<delivery_id>/acceptance.xls")
def acceptance(batch_id, delivery_id):
    model = _model_or_redirect(batch_id)
    delivery = model and next((d for d in model["deliveries"] if d["id"] == delivery_id), None)
    if not delivery:
        flash("Доставка не найдена - документы могли измениться", "error")
        return redirect(url_for("deliveries.index"))
    return send_file(deliveries.acceptance_xls(model["label"], delivery), as_attachment=True,
                     download_name=_file_name("Приёмка", model, delivery),
                     mimetype="application/vnd.ms-excel")


@deliveries_bp.route("/b/<int:batch_id>/<delivery_id>/invoice-total.xls")
def invoice_total(batch_id, delivery_id):
    model = _model_or_redirect(batch_id)
    delivery = model and next((d for d in model["deliveries"] if d["id"] == delivery_id), None)
    if not delivery or not delivery["boxes_found"]:
        flash("В доставке нет коробок из инвойсов - Invoice total не из чего собрать", "error")
        return redirect(url_for("deliveries.batch_view", batch_id=batch_id) if model
                        else url_for("deliveries.index"))
    return send_file(deliveries.invoice_total_xls(delivery), as_attachment=True,
                     download_name=_file_name("Invoice total", model, delivery),
                     mimetype="application/vnd.ms-excel")
