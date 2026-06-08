"""SQLite-хранилище для поставщиков, позиций и истории цен."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Iterator, Optional


DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "data" / "prices.db"


SCHEMA = """
CREATE TABLE IF NOT EXISTS suppliers (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    kind TEXT NOT NULL DEFAULT 'material',
    price_policy TEXT NOT NULL DEFAULT 'volatile',
    inn TEXT,
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY,
    supplier_id INTEGER NOT NULL REFERENCES suppliers(id) ON DELETE CASCADE,
    project_id INTEGER REFERENCES projects(id) ON DELETE SET NULL,
    invoice_number TEXT,
    invoice_date TEXT,
    total_without_vat REAL,
    total_with_vat REAL,
    source_file TEXT,
    raw_text TEXT,
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (date('now'))
);

CREATE INDEX IF NOT EXISTS idx_invoices_project ON invoices(project_id);
CREATE INDEX IF NOT EXISTS idx_invoices_supplier ON invoices(supplier_id);

CREATE TABLE IF NOT EXISTS invoice_items (
    id INTEGER PRIMARY KEY,
    invoice_id INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
    line_no INTEGER,
    name TEXT NOT NULL,
    article_supplier TEXT,
    article_manufacturer TEXT,
    unit TEXT,
    quantity REAL,
    unit_price REAL,
    vat_rate REAL,
    vat_included INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_invoice_items_invoice ON invoice_items(invoice_id);
CREATE INDEX IF NOT EXISTS idx_invoice_items_article_s ON invoice_items(article_supplier);
CREATE INDEX IF NOT EXISTS idx_invoice_items_article_m ON invoice_items(article_manufacturer);
"""


@contextmanager
def connect(db_path: Path = DEFAULT_DB_PATH) -> Iterator[sqlite3.Connection]:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(db_path: Path = DEFAULT_DB_PATH) -> None:
    with connect(db_path) as conn:
        conn.executescript(SCHEMA)


def upsert_supplier(
    name: str,
    kind: str = "material",
    price_policy: str = "volatile",
    notes: Optional[str] = None,
    db_path: Path = DEFAULT_DB_PATH,
) -> int:
    with connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO suppliers(name, kind, price_policy, notes) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(name) DO UPDATE SET kind=excluded.kind, "
            "price_policy=excluded.price_policy, notes=excluded.notes "
            "RETURNING id",
            (name, kind, price_policy, notes),
        )
        return cur.fetchone()["id"]


def upsert_item(
    name: str,
    article: Optional[str] = None,
    unit: Optional[str] = None,
    canonical_name: Optional[str] = None,
    db_path: Path = DEFAULT_DB_PATH,
) -> int:
    with connect(db_path) as conn:
        if article:
            row = conn.execute(
                "SELECT id FROM items WHERE article = ?", (article,)
            ).fetchone()
            if row:
                return row["id"]
        cur = conn.execute(
            "INSERT INTO items(name, article, unit, canonical_name) "
            "VALUES (?, ?, ?, ?) RETURNING id",
            (name, article, unit, canonical_name or name.lower()),
        )
        return cur.fetchone()["id"]


def add_price(
    supplier_id: int,
    item_id: int,
    price: float,
    vat_included: bool = True,
    quoted_on: Optional[str] = None,
    source_file: Optional[str] = None,
    notes: Optional[str] = None,
    db_path: Path = DEFAULT_DB_PATH,
) -> int:
    with connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO prices(supplier_id, item_id, price, vat_included, "
            "quoted_on, source_file, notes) VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id",
            (
                supplier_id,
                item_id,
                price,
                1 if vat_included else 0,
                quoted_on or date.today().isoformat(),
                source_file,
                notes,
            ),
        )
        return cur.fetchone()["id"]


def latest_price(item_id: int, supplier_id: Optional[int] = None,
                 db_path: Path = DEFAULT_DB_PATH) -> Optional[sqlite3.Row]:
    with connect(db_path) as conn:
        if supplier_id is not None:
            row = conn.execute(
                "SELECT * FROM prices WHERE item_id = ? AND supplier_id = ? "
                "ORDER BY quoted_on DESC, id DESC LIMIT 1",
                (item_id, supplier_id),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT * FROM prices WHERE item_id = ? "
                "ORDER BY quoted_on DESC, id DESC LIMIT 1",
                (item_id,),
            ).fetchone()
        return row


def list_suppliers(db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    with connect(db_path) as conn:
        return list(conn.execute("SELECT * FROM suppliers ORDER BY name"))


def price_history(item_id: int, db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    with connect(db_path) as conn:
        return list(
            conn.execute(
                "SELECT p.*, s.name AS supplier_name FROM prices p "
                "JOIN suppliers s ON s.id = p.supplier_id "
                "WHERE p.item_id = ? ORDER BY p.quoted_on DESC, p.id DESC",
                (item_id,),
            )
        )


# ------------------------------------------------------- Projects
def upsert_project(name: str, notes: Optional[str] = None,
                    db_path: Path = DEFAULT_DB_PATH) -> int:
    with connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO projects(name, notes) VALUES (?, ?) "
            "ON CONFLICT(name) DO UPDATE SET notes=COALESCE(excluded.notes, notes) "
            "RETURNING id",
            (name, notes),
        )
        return cur.fetchone()["id"]


def list_projects(db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    with connect(db_path) as conn:
        return list(conn.execute("SELECT * FROM projects ORDER BY name"))


def get_project_by_name(name: str,
                         db_path: Path = DEFAULT_DB_PATH) -> Optional[sqlite3.Row]:
    with connect(db_path) as conn:
        return conn.execute("SELECT * FROM projects WHERE name = ?", (name,)).fetchone()


# ------------------------------------------------------- Invoices
def save_invoice(
    supplier_id: int,
    project_id: Optional[int],
    invoice_number: Optional[str],
    invoice_date: Optional[str],
    total_without_vat: Optional[float],
    total_with_vat: Optional[float],
    items: list[dict],
    source_file: Optional[str] = None,
    raw_text: Optional[str] = None,
    notes: Optional[str] = None,
    db_path: Path = DEFAULT_DB_PATH,
) -> int:
    """Сохранить счёт целиком (заголовок + позиции) в одной транзакции."""
    with connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO invoices(supplier_id, project_id, invoice_number, "
            "invoice_date, total_without_vat, total_with_vat, source_file, "
            "raw_text, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id",
            (supplier_id, project_id, invoice_number, invoice_date,
             total_without_vat, total_with_vat, source_file, raw_text, notes),
        )
        invoice_id = cur.fetchone()["id"]
        for it in items:
            conn.execute(
                "INSERT INTO invoice_items(invoice_id, line_no, name, "
                "article_supplier, article_manufacturer, unit, quantity, "
                "unit_price, vat_rate, vat_included) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    invoice_id,
                    it.get("line_no"),
                    it.get("name", ""),
                    it.get("article_supplier"),
                    it.get("article_manufacturer"),
                    it.get("unit"),
                    it.get("quantity"),
                    it.get("unit_price"),
                    it.get("vat_rate"),
                    1 if it.get("vat_included") else 0,
                ),
            )
        return invoice_id


def list_invoices(project_id: Optional[int] = None,
                   db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    with connect(db_path) as conn:
        if project_id is not None:
            return list(conn.execute(
                "SELECT i.*, s.name AS supplier_name, p.name AS project_name "
                "FROM invoices i "
                "JOIN suppliers s ON s.id = i.supplier_id "
                "LEFT JOIN projects p ON p.id = i.project_id "
                "WHERE i.project_id = ? "
                "ORDER BY i.invoice_date DESC, i.id DESC",
                (project_id,),
            ))
        return list(conn.execute(
            "SELECT i.*, s.name AS supplier_name, p.name AS project_name "
            "FROM invoices i "
            "JOIN suppliers s ON s.id = i.supplier_id "
            "LEFT JOIN projects p ON p.id = i.project_id "
            "ORDER BY i.invoice_date DESC, i.id DESC",
        ))


def list_invoice_items(invoice_id: int,
                        db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    with connect(db_path) as conn:
        return list(conn.execute(
            "SELECT * FROM invoice_items WHERE invoice_id = ? ORDER BY line_no, id",
            (invoice_id,),
        ))


def delete_invoice(invoice_id: int, db_path: Path = DEFAULT_DB_PATH) -> None:
    with connect(db_path) as conn:
        conn.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))


def all_invoice_items_for_project(project_id: int,
                                   db_path: Path = DEFAULT_DB_PATH) -> list[sqlite3.Row]:
    """Все позиции счетов для проекта — нужно для AI-сопоставления."""
    with connect(db_path) as conn:
        return list(conn.execute(
            "SELECT ii.*, i.invoice_number, i.invoice_date, "
            "s.name AS supplier_name "
            "FROM invoice_items ii "
            "JOIN invoices i ON i.id = ii.invoice_id "
            "JOIN suppliers s ON s.id = i.supplier_id "
            "WHERE i.project_id = ?",
            (project_id,),
        ))
