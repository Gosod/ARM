# -*- coding: utf-8 -*-
import os
import time
import threading
from functools import wraps
from flask import Flask, request, jsonify, render_template, redirect, url_for, session

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

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
IMAGES_FOLDER = os.path.join(os.path.dirname(__file__), "static", "images")
PDFS_FOLDER   = os.path.join(os.path.dirname(__file__), "static", "pdfs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGES_FOLDER, exist_ok=True)
os.makedirs(PDFS_FOLDER,   exist_ok=True)

app = Flask(__name__)
app.secret_key = "fm-tracker-secret-2026"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

init_db()


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
    return {
        "id":       session.get("user_id"),
        "username": session.get("username"),
        "role":     session.get("role"),
        "shift":    session.get("shift"),
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
    path = os.path.join(os.path.dirname(__file__), "static", folder, filename)
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
        session["role"]     = user["role"]
        session["shift"]    = shift if user["role"] == "painting" else None

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
    raw_orders = get_all_orders()
    orders = []
    for o in raw_orders:
        orders.append({
            "id":           o["id"],
            "order_number": o["order_number"],
            "created_at":   o["created_at"],
            "complete":     is_order_complete(o["id"])
        })
    if user["role"] == "admin":
        return render_template("admin.html", orders=orders, user=user,
                               sheets_ok=sheets.is_configured())
    else:
        return render_template("index.html", orders=orders, user=user,
                               sheets_ok=sheets.is_configured())


# ── Admin: users ─────────────────────────────────────────────────────────────

@app.route("/admin/users")
@admin_required
def admin_users():
    users = get_all_users()
    return render_template("admin_users.html", users=users, user=current_user())


@app.route("/admin/users/create", methods=["POST"])
@admin_required
def admin_create_user():
    data = request.get_json(silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    role     = data.get("role", "worker")
    if not username or not password:
        return jsonify({"ok": False, "error": "Заполните все поля"}), 400
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

    conn = get_conn()
    order = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()
    conn.close()

    # Участок покраски — свой шаблон
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
    existing = get_order_by_number(order_number)
    if existing:
        os.remove(filepath)
        return jsonify({"error": f"Order {order_number} already loaded", "order_id": existing["id"]}), 409

    order = create_order(order_number)
    insert_positions(order["id"], positions)
    os.remove(filepath)

    return jsonify({
        "ok":              True,
        "order_id":        order["id"],
        "order_number":    order_number,
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

    current_done = get_done_qty(position_id)
    remaining = pos["qty"] - current_done
    if remaining <= 0:
        return jsonify({"ok": True, "total_done": current_done, "qty": pos["qty"], "complete": True})

    qty_done = min(qty_done, remaining)
    user = current_user()
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
    orders = get_all_orders()
    return jsonify([dict(o) for o in orders])


@app.route("/api/users")
@admin_required
def api_users():
    users = get_all_users()
    return jsonify([{"id":u["id"],"username":u["username"],"role":u["role"],"created_at":u["created_at"]} for u in users])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
