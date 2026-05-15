# -*- coding: utf-8 -*-
import sqlite3
import os
import hashlib

DB_PATH = os.path.join(os.path.dirname(__file__), "tracker.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            username    TEXT NOT NULL UNIQUE,
            password    TEXT NOT NULL,
            role        TEXT NOT NULL DEFAULT 'worker',
            shift       TEXT,
            created_at  TEXT DEFAULT (datetime('now', 'localtime'))
        );

        CREATE TABLE IF NOT EXISTS orders (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            order_number TEXT NOT NULL UNIQUE,
            section      TEXT,
            created_at   TEXT DEFAULT (datetime('now', 'localtime'))
        );

        CREATE TABLE IF NOT EXISTS positions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id    INTEGER NOT NULL REFERENCES orders(id),
            pos_number  INTEGER,
            designation TEXT,
            name        TEXT,
            qty         REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS markings (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            position_id INTEGER NOT NULL REFERENCES positions(id),
            qty_done    REAL NOT NULL,
            user_id     INTEGER REFERENCES users(id),
            shift       TEXT,
            marked_at   TEXT DEFAULT (datetime('now', 'localtime')),
            synced      INTEGER DEFAULT 0
        );
    """)

    # Миграции для существующих tracker.db без потери данных.
    order_cols = [r["name"] for r in conn.execute("PRAGMA table_info(orders)").fetchall()]
    if "section" not in order_cols:
        conn.execute("ALTER TABLE orders ADD COLUMN section TEXT")

    # Старую роль worker считаем участком механосборки, чтобы существующие
    # сотрудники не потеряли доступ после обновления ролей.
    conn.execute("UPDATE users SET role = 'assembly' WHERE role = 'worker'")

    # Создать дефолтного админа если нет ни одного пользователя
    existing = conn.execute("SELECT COUNT(*) as cnt FROM users").fetchone()
    if existing["cnt"] == 0:
        conn.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("admin", hash_password("admin"), "admin")
        )

    conn.commit()
    conn.close()


# ── Пользователи ─────────────────────────────────────────────────────────────

def get_user_by_username(username):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return row


def get_user_by_id(user_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    return row


def get_all_users():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM users ORDER BY role, username").fetchall()
    conn.close()
    return rows


def create_user(username, password, role):
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            (username, hash_password(password), role)
        )
        conn.commit()
        return True, None
    except sqlite3.IntegrityError:
        return False, "Пользователь уже существует"
    finally:
        conn.close()


def delete_user(user_id):
    conn = get_conn()
    conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


def verify_user(username, password):
    user = get_user_by_username(username)
    if user and user["password"] == hash_password(password):
        return user
    return None


# ── Заказы ───────────────────────────────────────────────────────────────────

def get_all_orders(section=None):
    conn = get_conn()
    if section:
        rows = conn.execute(
            "SELECT * FROM orders WHERE section = ? ORDER BY created_at DESC",
            (section,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM orders ORDER BY created_at DESC").fetchall()
    conn.close()
    return rows


def get_order_by_number(order_number):
    conn = get_conn()
    row = conn.execute("SELECT * FROM orders WHERE order_number = ?", (order_number,)).fetchone()
    conn.close()
    return row


def create_order(order_number, section=None):
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO orders (order_number, section) VALUES (?, ?)",
            (order_number, section)
        )
        conn.commit()
        row = conn.execute("SELECT * FROM orders WHERE order_number = ?", (order_number,)).fetchone()
        return row
    finally:
        conn.close()


def delete_order(order_id):
    conn = get_conn()
    conn.execute("DELETE FROM markings WHERE position_id IN (SELECT id FROM positions WHERE order_id = ?)", (order_id,))
    conn.execute("DELETE FROM positions WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()


# ── Позиции ──────────────────────────────────────────────────────────────────

def get_positions(order_id):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM positions WHERE order_id = ? ORDER BY pos_number", (order_id,)).fetchall()
    conn.close()
    return rows


def insert_positions(order_id, positions):
    conn = get_conn()
    conn.executemany(
        "INSERT INTO positions (order_id, pos_number, designation, name, qty) VALUES (:order_id, :pos_number, :designation, :name, :qty)",
        [{"order_id": order_id, **p} for p in positions]
    )
    conn.commit()
    conn.close()


def get_position(position_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM positions WHERE id = ?", (position_id,)).fetchone()
    conn.close()
    return row


# ── Отметки ──────────────────────────────────────────────────────────────────

def add_marking(position_id, qty_done, user_id=None, shift=None):
    conn = get_conn()
    conn.execute(
        "INSERT INTO markings (position_id, qty_done, user_id, shift) VALUES (?, ?, ?, ?)",
        (position_id, qty_done, user_id, shift)
    )
    conn.commit()
    row = conn.execute(
        "SELECT * FROM markings WHERE position_id = ? ORDER BY id DESC LIMIT 1",
        (position_id,)
    ).fetchone()
    conn.close()
    return row


def get_unsynced_markings():
    conn = get_conn()
    rows = conn.execute("""
        SELECT m.*, p.designation, p.name, p.qty, o.order_number
        FROM markings m
        JOIN positions p ON m.position_id = p.id
        JOIN orders o ON p.order_id = o.id
        WHERE m.synced = 0
    """).fetchall()
    conn.close()
    return rows


def mark_synced(marking_id):
    conn = get_conn()
    conn.execute("UPDATE markings SET synced = 1 WHERE id = ?", (marking_id,))
    conn.commit()
    conn.close()


def get_done_qty(position_id):
    conn = get_conn()
    row = conn.execute(
        "SELECT COALESCE(SUM(qty_done), 0) as total FROM markings WHERE position_id = ?",
        (position_id,)
    ).fetchone()
    conn.close()
    return row["total"] if row else 0


def is_order_complete(order_id):
    """Возвращает True если все позиции заказа выполнены."""
    conn = get_conn()
    positions = conn.execute(
        "SELECT id, qty FROM positions WHERE order_id = ?", (order_id,)
    ).fetchall()
    conn.close()
    if not positions:
        return False
    for p in positions:
        done = get_done_qty(p["id"])
        if done < p["qty"]:
            return False
    return True
