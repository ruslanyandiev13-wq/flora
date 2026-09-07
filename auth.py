"""
Полноценные аккаунты пользователей с ролями (admin/user), т.к. до этого в
приложении не было вообще никакой внутренней системы пользователей - только
общий Basic Auth на уровне nginx (см. ТЗ "Биллинг / История", раздел 6.1:
проверка роли должна быть на уровне роута, а не скрытием кнопки).

Nginx Basic Auth можно оставить как есть (внешний защитный периметр) - этот
модуль добавляет ВТОРОЙ, внутренний уровень: кто именно из пользователей
сейчас работает и какая у него роль, чтобы billing и голландский/импортный
модули знали настоящего uploaded_by и могли ограничивать admin-функции.
"""
import functools
import os
import secrets

from flask import Blueprint, request, render_template, redirect, url_for, session, flash
from werkzeug.security import generate_password_hash, check_password_hash

import db

auth_bp = Blueprint("auth", __name__)

# Роуты, доступные без входа в систему - логин и статика.
_PUBLIC_ENDPOINTS = {"auth.login", "static"}


def init_db():
    conn = db.get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK (role IN ('admin', 'user')),
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()


def create_user(username, password, role):
    if role not in ("admin", "user"):
        raise ValueError("role должна быть 'admin' или 'user'")
    conn = db.get_conn()
    conn.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                 (username, generate_password_hash(password), role))
    conn.commit()
    conn.close()


def list_users():
    conn = db.get_conn()
    rows = conn.execute("SELECT id, username, role, created_at FROM users ORDER BY username").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_user(user_id):
    conn = db.get_conn()
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()


def _verify(username, password):
    conn = db.get_conn()
    row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    if row and check_password_hash(row["password_hash"], password):
        return dict(row)
    return None


def current_user():
    if "user_id" not in session:
        return None
    return {"id": session["user_id"], "username": session["username"], "role": session["role"]}


def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if current_user() is None:
            return redirect(url_for("auth.login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        user = current_user()
        if user is None:
            return redirect(url_for("auth.login", next=request.path))
        if user["role"] != "admin":
            flash("Эта страница доступна только администратору", "error")
            return redirect(url_for("index"))
        return view(*args, **kwargs)
    return wrapped


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = _verify(request.form.get("username", ""), request.form.get("password", ""))
        if user:
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            next_url = request.args.get("next") or url_for("index")
            return redirect(next_url)
        flash("Неверный логин или пароль", "error")
    return render_template("login.html")


@auth_bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))


def bootstrap_admin():
    """Если пользователей вообще ещё нет - создаёт первого admin'а, чтобы
    было кем войти и создать остальных через /billing/admin/users. Пароль -
    из ADMIN_PASSWORD (см. .env, по аналогии с остальными секретами
    деплоя), если не задан - генерируется и печатается в лог один раз."""
    if list_users():
        return
    username = os.environ.get("ADMIN_USERNAME", "admin")
    password = os.environ.get("ADMIN_PASSWORD")
    generated = password is None
    if generated:
        password = secrets.token_urlsafe(12)
    create_user(username, password, "admin")
    if generated:
        print(f"[auth] Создан первый администратор: {username} / {password} "
              f"(сохраните пароль - он больше нигде не показывается; "
              f"чтобы задать свой, установите ADMIN_PASSWORD в .env и пересоздайте пользователя)")


def init_app(app):
    init_db()
    bootstrap_admin()
    app.register_blueprint(auth_bp)
    app.context_processor(lambda: {"current_user": current_user()})

    @app.before_request
    def _require_login():
        if request.endpoint is None or request.endpoint in _PUBLIC_ENDPOINTS:
            return None
        if current_user() is None:
            return redirect(url_for("auth.login", next=request.path))
        return None
