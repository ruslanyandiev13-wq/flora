"""
Постоянное хранилище исходных файлов и результатов по каждой накладной
(invoices.id). Нужно для чат-правок (см. ТЗ "Чат-помощник и обратная связь",
раздел 1) - в отличие от временных каталогов загрузки (tempfile.mkdtemp()),
эти копии не чистятся и переживают рестарт сервиса, поэтому "тяжёлую" правку
можно применить даже к накладной недельной давности.
"""
import os
import shutil

DOCS_DIR = os.path.join(os.path.dirname(__file__), "data", "documents")


def _doc_dir(invoice_id):
    d = os.path.join(DOCS_DIR, str(invoice_id))
    os.makedirs(d, exist_ok=True)
    return d


def store_source(invoice_id, src_path):
    """Копирует исходный загруженный файл в постоянное хранилище, возвращает
    новый путь. Возвращает None, если исходника уже нет (temp-каталог мог
    быть очищен раньше, чем успели прикрепить документ)."""
    if not src_path or not os.path.exists(src_path):
        return None
    ext = os.path.splitext(src_path)[1]
    dest = os.path.join(_doc_dir(invoice_id), f"source{ext}")
    shutil.copy2(src_path, dest)
    return dest


def store_result(invoice_id, src_path):
    if not src_path or not os.path.exists(src_path):
        return None
    ext = os.path.splitext(src_path)[1]
    dest = os.path.join(_doc_dir(invoice_id), f"result{ext}")
    shutil.copy2(src_path, dest)
    return dest
