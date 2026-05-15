# -*- coding: utf-8 -*-
import os
import re
import time
import threading
from functools import wraps
from flask import Flask, request, jsonify, render_template, redirect, url_for, session
from jinja2 import ChoiceLoader, FileSystemLoader

from db import (
    is_order_complete,
    init_db, get_all_orders, get_order_by_number, create_order,
    delete_order, get_positions, insert_positions, get_position,
    add_marking, get_unsynced_markings, mark_synced, get_done_qty,
    get_conn, verify_user, get_all_users, create_user, delete_user,
    get_user_by_id
)
from excel_parser import parse_excel
import sheets

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
IMAGES_FOLDER = os.path.join(BASE_DIR, "static", "images")
PDFS_FOLDER   = os.path.join(BASE_DIR, "static", "pdfs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGES_FOLDER, exist_ok=True)
os.makedirs(PDFS_FOLDER,   exist_ok=True)

app = Flask(__name__)
app.jinja_loader = ChoiceLoader([
    FileSystemLoader(BASE_DIR),
    FileSystemLoader(os.path.join(BASE_DIR, "templates")),
])
app.secret_key = "fm-tracker-secret-2026"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

init_db()

ROLE_TITLES = {
    "admin": "Администратор",
    "assembly": "Участок механосборки",
    "welding": "Участок слесарно-сварочный",
    "painting": "Участок покраски",
    "quality": "Отдел технического контроля",
    "storekeeper": "Склад",
}
PRODUCTION_ROLES = {"assembly", "welding", "painting"}
SERVICE_ROLES = {"quality", "storekeeper"}
ALLOWED_ROLES = {"admin", *PRODUCTION_ROLES, *SERVICE_ROLES}


def normalize_role(role):
    if role == "worker":
        return "assembly"
    return role


def role_title(role):
    return ROLE_TITLES.get(normalize_role(role), role or "")


def detect_section(text):
    low = (text or "").lower()
    if "покрас" in low or "paint" in low:
        return "painting"
    if "слесар" in low or "свар" in low or "сл-св" in low or "weld" in low:
        return "welding"
    if "механо" in low or "механ" in low or "мех" in low or "механосбор" in low or "сбор" in low or "assembl" in low:
        return "assembly"
    return None


def section_from_upload(filename, parsed_order_number=None):
    # Участок ожидается в скобках имени Excel-файла, например:
    #   Заказ-123 (покраска).xlsx
    # Дополнительно смотрим номер заказа из файла как запасной вариант.
    base = os.path.splitext(os.path.basename(filename or ""))[0]
    for candidate in re.findall(r"\(([^)]*)\)", base):
        section = detect_section(candidate)
        if section:
            return section
    return detect_section(base) or detect_section(parsed_order_number)


def backfill_order_sections():
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT id, order_number FROM orders WHERE section IS NULL OR section = ''"
        ).fetchall()
        for row in rows:
            section = section_from_upload(row["order_number"])
            if section:
                conn.execute("UPDATE orders SET section = ? WHERE id = ?", (section, row["id"]))
        conn.commit()
    finally:
        conn.close()


backfill_order_sections()

# ── Auth helpers ─────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated


def current_user():
    role = normalize_role(session.get("role"))
    return {
        "id":            session.get("user_id"),
        "username":      session.get("username"),
        "role":          role,
        "shift":         session.get("shift"),
        "section_title": role_title(role),
    }


# ── Sheets sync ──────────────────────────────────────────────────────────────

def sync_to_sheets():
    if not sheets.is_configured():
        return
    unsynced = get_unsynced_markings()
    if not unsynced:
        return
    positions_to_sync = {}
    for m in unsynced:
        pid = m["position_id"]
        if pid not in positions_to_sync:
            positions_to_sync[pid] = m
    for pid, m in positions_to_sync.items():
        try:
            total_done = get_done_qty(pid)
            sheets.upsert_position(
                order_number=m["order_number"],
                designation=m["designation"],
                name=m["name"],
                total_qty_done=total_done
            )
            conn = get_conn()
            conn.execute("UPDATE markings SET synced = 1 WHERE position_id = ? AND synced = 0", (pid,))
            conn.commit()
            conn.close()
            time.sleep(1)
        except Exception as e:
            app.logger.error(f"Sync error for position {pid}: {e}")
            break


def sync_background():
    t = threading.Thread(target=sync_to_sheets, daemon=True)
    t.start()


# ── Asset helper ─────────────────────────────────────────────────────────────

def asset_url(designation, folder, ext):
    if not designation:
        return None
    filename = designation.strip() + ext
    path = os.path.join(BASE_DIR, "static", folder, filename)
    if os.path.exists(path):
        return f"/static/{folder}/{filename}"
    return None


# ── Auth routes ──────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        mode = data.get("mode", "password")

        if mode == "guest":
            session["user_id"]  = 0
            session["username"] = "Гостевой режим"
            session["role"]     = "guest"
            session["shift"]    = None
            return jsonify({"ok": True, "redirect": url_for("dashboard")})

        username = data.get("username", "").strip()
        password = data.get("password", "")
        shift    = data.get("shift")

        user = verify_user(username, password)
        if not user:
            return jsonify({"ok": False, "error": "Неверный логин или пароль"}), 401

        session["user_id"]  = user["id"]
        session["username"] = user["username"]
        role = normalize_role(user["role"])
        session["role"]     = role
        session["shift"]    = shift if role == "painting" else None

        return jsonify({"ok": True, "redirect": url_for("dashboard")})

    # GET
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Main dashboard ────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def dashboard():
    user = current_user()
    if user["role"] == "admin":
        raw_orders = get_all_orders()
    elif user["role"] in PRODUCTION_ROLES:
        raw_orders = get_all_orders(user["role"])
    else:
        # Для ОТК и склада логика очередей/уведомлений будет добавлена отдельно.
        raw_orders = []
    orders = []
    for o in raw_orders:
        orders.append({
            "id":            o["id"],
            "order_number":  o["order_number"],
            "section":       o["section"],
            "section_title": role_title(o["section"]),
            "created_at":    o["created_at"],
            "complete":      is_order_complete(o["id"])
        })
    if user["role"] == "admin":
        return render_template("admin.html", orders=orders, user=user,
                               sheets_ok=sheets.is_configured())
    if user["role"] == "quality":
        return render_template("quality.html", orders=orders, user=user,
                               sheets_ok=sheets.is_configured())
    if user["role"] == "storekeeper":
        return render_template("storekeeper.html", orders=orders, user=user,
                               sheets_ok=sheets.is_configured())
    return render_template("index.html", orders=orders, user=user,
                           sheets_ok=sheets.is_configured())


# ── Admin: users ─────────────────────────────────────────────────────────────

@app.route("/admin/users")
@admin_required
def admin_users():
    return redirect(url_for("dashboard"))


@app.route("/admin/users/create", methods=["POST"])
@admin_required
def admin_create_user():
    data = request.get_json(silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    role     = normalize_role(data.get("role", "assembly"))
    if not username or not password:
        return jsonify({"ok": False, "error": "Заполните все поля"}), 400
    if role not in ALLOWED_ROLES:
        return jsonify({"ok": False, "error": "Неизвестная роль"}), 400
    ok, err = create_user(username, password, role)
    if not ok:
        return jsonify({"ok": False, "error": err}), 409
    return jsonify({"ok": True})


@app.route("/admin/users/delete/<int:user_id>", methods=["POST"])
@admin_required
def admin_delete_user(user_id):
    if user_id == session.get("user_id"):
        return jsonify({"ok": False, "error": "Нельзя удалить себя"}), 400
    delete_user(user_id)
    return jsonify({"ok": True})


# ── Order routes ─────────────────────────────────────────────────────────────

@app.route("/order/<int:order_id>")
@login_required
def order_view(order_id):
    user = current_user()
    conn = get_conn()
    order = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()
    conn.close()
    if not order:
        return redirect(url_for("dashboard"))
    if user["role"] != "admin" and order["section"] != user["role"]:
        return redirect(url_for("dashboard"))

    positions = get_positions(order_id)
    if not positions:
        return redirect(url_for("dashboard"))

    pos_list = []
    for p in positions:
        done = get_done_qty(p["id"])
        pos_list.append({
            "id":          p["id"],
            "pos_number":  p["pos_number"],
            "designation": p["designation"],
            "name":        p["name"],
            "qty":         p["qty"],
            "done":        done,
            "complete":    done >= p["qty"] and p["qty"] > 0,
            "image_url":   asset_url(p["designation"], "images", ".jpg"),
            "pdf_url":     asset_url(p["designation"], "pdfs",   ".pdf"),
        })

    # Участок покраски — свой шаблон с тем же функционалом, но отдельной визуальной настройкой.
    template = "order_painting.html" if user["role"] == "painting" else "order.html"
    return render_template(template, order=order, positions=pos_list, user=user)


@app.route("/upload", methods=["POST"])
@admin_required
def upload():
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files selected"}), 400

    excel_file = None
    images, pdfs = [], []

    for f in files:
        low = f.filename.lower()
        if low.endswith((".xlsx", ".xls")):
            excel_file = f
        elif low.endswith((".jpg", ".jpeg", ".png")):
            images.append(f)
        elif low.endswith(".pdf"):
            pdfs.append(f)

    if not excel_file:
        return jsonify({"error": "Excel file not found"}), 400

    for img in images:
        img.save(os.path.join(IMAGES_FOLDER, os.path.basename(img.filename)))
    for pdf in pdfs:
        pdf.save(os.path.join(PDFS_FOLDER, os.path.basename(pdf.filename)))

    order_number = os.path.splitext(os.path.basename(excel_file.filename))[0]
    filepath = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(filepath)

    try:
        result = parse_excel(filepath)
    except ValueError as e:
        os.remove(filepath)
        return jsonify({"error": str(e)}), 422

    positions = result["positions"]
    section = section_from_upload(excel_file.filename, result.get("order_number"))
    if not section:
        os.remove(filepath)
        return jsonify({
            "error": "Не удалось определить участок. Добавьте участок в скобках имени Excel-файла: (механосборка), (слесарно-сварочный) или (покраска)"
        }), 422

    existing = get_order_by_number(order_number)
    if existing:
        os.remove(filepath)
        return jsonify({"error": f"Order {order_number} already loaded", "order_id": existing["id"]}), 409

    order = create_order(order_number, section)
    insert_positions(order["id"], positions)
    os.remove(filepath)

    return jsonify({
        "ok":              True,
        "order_id":        order["id"],
        "order_number":    order_number,
        "section":         section,
        "section_title":   role_title(section),
        "positions_count": len(positions),
        "images_saved":    len(images),
        "pdfs_saved":      len(pdfs),
    })


@app.route("/mark/<int:position_id>", methods=["POST"])
@login_required
def mark(position_id):
    data = request.get_json(silent=True) or {}
    qty_done = data.get("qty_done", 1)

    try:
        qty_done = float(qty_done)
    except (ValueError, TypeError):
        return jsonify({"error": "Invalid quantity"}), 400

    pos = get_position(position_id)
    if not pos:
        return jsonify({"error": "Position not found"}), 404

    user = current_user()
    conn = get_conn()
    order = conn.execute("SELECT section FROM orders WHERE id = ?", (pos["order_id"],)).fetchone()
    conn.close()
    if user["role"] != "admin" and (not order or order["section"] != user["role"]):
        return jsonify({"error": "Нет доступа к заказу другого участка"}), 403

    current_done = get_done_qty(position_id)
    remaining = pos["qty"] - current_done
    if remaining <= 0:
        return jsonify({"ok": True, "total_done": current_done, "qty": pos["qty"], "complete": True})

    qty_done = min(qty_done, remaining)
    add_marking(position_id, qty_done, user_id=user["id"] or None, shift=user["shift"])
    total_done = get_done_qty(position_id)
    sync_background()

    return jsonify({
        "ok":        True,
        "total_done": total_done,
        "qty":        pos["qty"],
        "complete":   total_done >= pos["qty"],
    })


@app.route("/delete/<int:order_id>", methods=["POST"])
@admin_required
def delete(order_id):
    delete_order(order_id)
    return jsonify({"ok": True})


@app.route("/sync")
@login_required
def manual_sync():
    sync_to_sheets()
    return jsonify({"ok": True})


@app.route("/api/orders")
@login_required
def api_orders():
    user = current_user()
    if user["role"] == "admin":
        orders = get_all_orders()
    elif user["role"] in PRODUCTION_ROLES:
        orders = get_all_orders(user["role"])
    else:
        orders = []
    return jsonify([dict(o) for o in orders])


@app.route("/api/users")
@admin_required
def api_users():
    users = get_all_users()
    return jsonify([
        {
            "id": u["id"],
            "username": u["username"],
            "role": normalize_role(u["role"]),
            "role_title": role_title(u["role"]),
            "created_at": u["created_at"],
        }
        for u in users
    ])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
