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
    kind TEXT NOT NULL DEFAULT 'material',   -- 'material' | 'work'
    price_policy TEXT NOT NULL DEFAULT 'volatile',  -- 'fixed' | 'volatile'
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    article TEXT,
    unit TEXT,
    canonical_name TEXT,
    created_at TEXT NOT NULL DEFAULT (date('now'))
);

CREATE INDEX IF NOT EXISTS idx_items_article ON items(article);
CREATE INDEX IF NOT EXISTS idx_items_canonical ON items(canonical_name);

CREATE TABLE IF NOT EXISTS prices (
    id INTEGER PRIMARY KEY,
    supplier_id INTEGER NOT NULL REFERENCES suppliers(id) ON DELETE CASCADE,
    item_id INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
    price REAL NOT NULL,
    currency TEXT NOT NULL DEFAULT 'RUB',
    vat_included INTEGER NOT NULL DEFAULT 1,
    quoted_on TEXT NOT NULL DEFAULT (date('now')),
    source_file TEXT,
    notes TEXT
);

CREATE INDEX IF NOT EXISTS idx_prices_item ON prices(item_id);
CREATE INDEX IF NOT EXISTS idx_prices_supplier ON prices(supplier_id);
CREATE INDEX IF NOT EXISTS idx_prices_date ON prices(quoted_on);
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
