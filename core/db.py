"""Хранилище для поставщиков, проектов, счетов и истории цен.

Работает на двух СУБД через SQLAlchemy Core:
- **Postgres** (production / Streamlit Cloud) — если задан `DATABASE_URL`
- **SQLite** (локальная разработка) — fallback на `data/prices.db`

Публичные функции возвращают `dict` (или `list[dict]`) — это совместимо как с
sqlite3.Row, так и с SQLAlchemy mappings: можно обращаться по ключу `row["id"]`.
"""
from __future__ import annotations

import os
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Any, Iterator, Optional

from sqlalchemy import (
    Boolean,
    Column,
    Date,
    Float,
    ForeignKey,
    Index,
    Integer,
    MetaData,
    String,
    Table,
    Text,
    create_engine,
    delete,
    select,
    text,
)
from sqlalchemy.dialects.postgresql import insert as pg_insert
from sqlalchemy.dialects.sqlite import insert as sqlite_insert
from sqlalchemy.engine import Engine


# --------------------------------------------------------------- Engine
DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "data" / "prices.db"


def _resolve_database_url() -> str:
    """Определяем URL подключения.

    Приоритет:
    1. ENV `DATABASE_URL` (production)
    2. ENV `SMETA_DB_URL` (альтернатива)
    3. SQLite по DEFAULT_DB_PATH (локальная разработка)
    """
    url = os.getenv("DATABASE_URL") or os.getenv("SMETA_DB_URL")
    if url:
        # Neon/Heroku/Render иногда дают URL вида postgres://... — SQLAlchemy
        # ждёт postgresql+psycopg://
        if url.startswith("postgres://"):
            url = "postgresql+psycopg://" + url[len("postgres://"):]
        elif url.startswith("postgresql://") and "+psycopg" not in url:
            url = "postgresql+psycopg://" + url[len("postgresql://"):]
        return url
    DEFAULT_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{DEFAULT_DB_PATH}"


_engine: Optional[Engine] = None


def get_engine() -> Engine:
    global _engine
    if _engine is None:
        url = _resolve_database_url()
        connect_args: dict[str, Any] = {}
        if url.startswith("sqlite"):
            connect_args["check_same_thread"] = False
        _engine = create_engine(url, future=True, pool_pre_ping=True,
                                  connect_args=connect_args)
    return _engine


def is_postgres() -> bool:
    return get_engine().dialect.name == "postgresql"


def db_label() -> str:
    """Короткая метка для отображения в UI."""
    eng = get_engine()
    if eng.dialect.name == "postgresql":
        host = eng.url.host or "?"
        return f"Postgres @ {host}"
    return f"SQLite ({Path(eng.url.database or '').name})"


# --------------------------------------------------------------- Schema
metadata = MetaData()


suppliers = Table(
    "suppliers", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("name", String(200), nullable=False, unique=True),
    Column("kind", String(20), nullable=False, server_default=text("'material'")),
    Column("price_policy", String(20), nullable=False,
            server_default=text("'volatile'")),
    Column("inn", String(20)),
    Column("notes", Text),
    Column("created_at", Date, nullable=False,
            server_default=text("CURRENT_DATE")),
)


projects = Table(
    "projects", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("name", String(200), nullable=False, unique=True),
    Column("notes", Text),
    Column("created_at", Date, nullable=False,
            server_default=text("CURRENT_DATE")),
)


items = Table(
    "items", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("name", String(500), nullable=False),
    Column("article", String(100)),
    Column("unit", String(20)),
    Column("canonical_name", String(500)),
    Column("created_at", Date, nullable=False,
            server_default=text("CURRENT_DATE")),
    Index("idx_items_article", "article"),
    Index("idx_items_canonical", "canonical_name"),
)


prices = Table(
    "prices", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("supplier_id", Integer,
            ForeignKey("suppliers.id", ondelete="CASCADE"), nullable=False),
    Column("item_id", Integer,
            ForeignKey("items.id", ondelete="CASCADE"), nullable=False),
    Column("price", Float, nullable=False),
    Column("currency", String(8), nullable=False, server_default=text("'RUB'")),
    Column("vat_included", Boolean, nullable=False, server_default=text("TRUE")),
    Column("quoted_on", Date, nullable=False, server_default=text("CURRENT_DATE")),
    Column("source_file", String(500)),
    Column("notes", Text),
    Index("idx_prices_item", "item_id"),
    Index("idx_prices_supplier", "supplier_id"),
    Index("idx_prices_date", "quoted_on"),
)


invoices = Table(
    "invoices", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("supplier_id", Integer,
            ForeignKey("suppliers.id", ondelete="CASCADE"), nullable=False),
    Column("project_id", Integer,
            ForeignKey("projects.id", ondelete="SET NULL")),
    Column("invoice_number", String(100)),
    Column("invoice_date", String(20)),
    Column("total_without_vat", Float),
    Column("total_with_vat", Float),
    Column("source_file", String(500)),
    Column("raw_text", Text),
    Column("notes", Text),
    Column("created_at", Date, nullable=False,
            server_default=text("CURRENT_DATE")),
    Index("idx_invoices_project", "project_id"),
    Index("idx_invoices_supplier", "supplier_id"),
)


invoice_items = Table(
    "invoice_items", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("invoice_id", Integer,
            ForeignKey("invoices.id", ondelete="CASCADE"), nullable=False),
    Column("line_no", Integer),
    Column("name", String(1000), nullable=False),
    Column("article_supplier", String(100)),
    Column("article_manufacturer", String(100)),
    Column("unit", String(20)),
    Column("quantity", Float),
    Column("unit_price", Float),
    Column("vat_rate", Float),
    Column("vat_included", Boolean, nullable=False,
            server_default=text("FALSE")),
    Index("idx_invoice_items_invoice", "invoice_id"),
    Index("idx_invoice_items_article_s", "article_supplier"),
    Index("idx_invoice_items_article_m", "article_manufacturer"),
)


def init_db(db_path: Optional[Path] = None) -> None:
    """Создаёт схему, если её нет. db_path сохранён для обратной совместимости."""
    if db_path is not None and not os.getenv("DATABASE_URL") \
            and not os.getenv("SMETA_DB_URL"):
        # Локальный SQLite на нестандартный путь
        global _engine
        _engine = None
        os.environ.pop("DATABASE_URL", None)
        # Подменяем DEFAULT_DB_PATH временно? Проще: если задан явный db_path —
        # создаём engine на него.
        db_path.parent.mkdir(parents=True, exist_ok=True)
        _engine = create_engine(
            f"sqlite:///{db_path}", future=True,
            connect_args={"check_same_thread": False},
        )
    metadata.create_all(get_engine())


# ------------------------------------------------------- Connection
@contextmanager
def connect(db_path: Optional[Path] = None) -> Iterator[Any]:
    """Контекстный менеджер для legacy-кода. Возвращает SA Connection."""
    eng = get_engine()
    with eng.begin() as conn:
        yield conn


def _rows_to_dicts(rows) -> list[dict]:
    return [dict(r._mapping) for r in rows]


def _row_to_dict(row) -> Optional[dict]:
    if row is None:
        return None
    return dict(row._mapping)


# ------------------------------------------------------- Suppliers
def upsert_supplier(
    name: str,
    kind: str = "material",
    price_policy: str = "volatile",
    notes: Optional[str] = None,
    db_path: Optional[Path] = None,
) -> int:
    with get_engine().begin() as conn:
        ins_cls = pg_insert if is_postgres() else sqlite_insert
        stmt = ins_cls(suppliers).values(
            name=name, kind=kind, price_policy=price_policy, notes=notes,
        )
        stmt = stmt.on_conflict_do_update(
            index_elements=[suppliers.c.name],
            set_={
                "kind": stmt.excluded.kind,
                "price_policy": stmt.excluded.price_policy,
                "notes": stmt.excluded.notes,
            },
        ).returning(suppliers.c.id)
        return conn.execute(stmt).scalar_one()


def list_suppliers(db_path: Optional[Path] = None) -> list[dict]:
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(suppliers).order_by(suppliers.c.name)
        ).fetchall()
        return _rows_to_dicts(rows)


# ------------------------------------------------------- Items / prices
def upsert_item(
    name: str,
    article: Optional[str] = None,
    unit: Optional[str] = None,
    canonical_name: Optional[str] = None,
    db_path: Optional[Path] = None,
) -> int:
    with get_engine().begin() as conn:
        if article:
            row = conn.execute(
                select(items.c.id).where(items.c.article == article)
            ).first()
            if row:
                return row[0]
        stmt = items.insert().values(
            name=name, article=article, unit=unit,
            canonical_name=canonical_name or name.lower(),
        ).returning(items.c.id)
        return conn.execute(stmt).scalar_one()


def add_price(
    supplier_id: int,
    item_id: int,
    price: float,
    vat_included: bool = True,
    quoted_on: Optional[str] = None,
    source_file: Optional[str] = None,
    notes: Optional[str] = None,
    db_path: Optional[Path] = None,
) -> int:
    with get_engine().begin() as conn:
        stmt = prices.insert().values(
            supplier_id=supplier_id,
            item_id=item_id,
            price=price,
            vat_included=vat_included,
            quoted_on=quoted_on or date.today().isoformat(),
            source_file=source_file,
            notes=notes,
        ).returning(prices.c.id)
        return conn.execute(stmt).scalar_one()


def latest_price(item_id: int, supplier_id: Optional[int] = None,
                  db_path: Optional[Path] = None) -> Optional[dict]:
    with get_engine().connect() as conn:
        q = select(prices).where(prices.c.item_id == item_id)
        if supplier_id is not None:
            q = q.where(prices.c.supplier_id == supplier_id)
        q = q.order_by(prices.c.quoted_on.desc(), prices.c.id.desc()).limit(1)
        return _row_to_dict(conn.execute(q).first())


def price_history(item_id: int, db_path: Optional[Path] = None) -> list[dict]:
    sql = text(
        "SELECT p.*, s.name AS supplier_name FROM prices p "
        "JOIN suppliers s ON s.id = p.supplier_id "
        "WHERE p.item_id = :item_id "
        "ORDER BY p.quoted_on DESC, p.id DESC"
    )
    with get_engine().connect() as conn:
        return _rows_to_dicts(conn.execute(sql, {"item_id": item_id}).fetchall())


# ------------------------------------------------------- Projects
def upsert_project(name: str, notes: Optional[str] = None,
                    db_path: Optional[Path] = None) -> int:
    with get_engine().begin() as conn:
        ins_cls = pg_insert if is_postgres() else sqlite_insert
        stmt = ins_cls(projects).values(name=name, notes=notes)
        # Сохраняем существующие notes если новое значение NULL
        stmt = stmt.on_conflict_do_update(
            index_elements=[projects.c.name],
            set_={"notes": text("COALESCE(EXCLUDED.notes, projects.notes)")},
        ).returning(projects.c.id)
        return conn.execute(stmt).scalar_one()


def list_projects(db_path: Optional[Path] = None) -> list[dict]:
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(projects).order_by(projects.c.name)
        ).fetchall()
        return _rows_to_dicts(rows)


def get_project_by_name(name: str,
                         db_path: Optional[Path] = None) -> Optional[dict]:
    with get_engine().connect() as conn:
        row = conn.execute(
            select(projects).where(projects.c.name == name)
        ).first()
        return _row_to_dict(row)


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
    db_path: Optional[Path] = None,
) -> int:
    with get_engine().begin() as conn:
        stmt = invoices.insert().values(
            supplier_id=supplier_id,
            project_id=project_id,
            invoice_number=invoice_number,
            invoice_date=invoice_date,
            total_without_vat=total_without_vat,
            total_with_vat=total_with_vat,
            source_file=source_file,
            raw_text=raw_text,
            notes=notes,
        ).returning(invoices.c.id)
        invoice_id = conn.execute(stmt).scalar_one()

        if items:
            conn.execute(
                invoice_items.insert(),
                [
                    {
                        "invoice_id": invoice_id,
                        "line_no": it.get("line_no"),
                        "name": it.get("name", ""),
                        "article_supplier": it.get("article_supplier"),
                        "article_manufacturer": it.get("article_manufacturer"),
                        "unit": it.get("unit"),
                        "quantity": it.get("quantity"),
                        "unit_price": it.get("unit_price"),
                        "vat_rate": it.get("vat_rate"),
                        "vat_included": bool(it.get("vat_included")),
                    }
                    for it in items
                ],
            )
        return invoice_id


def list_invoices(project_id: Optional[int] = None,
                   db_path: Optional[Path] = None) -> list[dict]:
    base = (
        "SELECT i.*, s.name AS supplier_name, p.name AS project_name "
        "FROM invoices i "
        "JOIN suppliers s ON s.id = i.supplier_id "
        "LEFT JOIN projects p ON p.id = i.project_id "
    )
    if project_id is not None:
        sql = text(base + "WHERE i.project_id = :pid "
                          "ORDER BY i.invoice_date DESC, i.id DESC")
        params = {"pid": project_id}
    else:
        sql = text(base + "ORDER BY i.invoice_date DESC, i.id DESC")
        params = {}
    with get_engine().connect() as conn:
        return _rows_to_dicts(conn.execute(sql, params).fetchall())


def list_invoice_items(invoice_id: int,
                        db_path: Optional[Path] = None) -> list[dict]:
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(invoice_items)
            .where(invoice_items.c.invoice_id == invoice_id)
            .order_by(invoice_items.c.line_no, invoice_items.c.id)
        ).fetchall()
        return _rows_to_dicts(rows)


def delete_invoice(invoice_id: int, db_path: Optional[Path] = None) -> None:
    with get_engine().begin() as conn:
        conn.execute(delete(invoices).where(invoices.c.id == invoice_id))


def all_invoice_items_for_project(project_id: int,
                                   db_path: Optional[Path] = None) -> list[dict]:
    """Все позиции счетов для проекта — нужно для AI-сопоставления."""
    sql = text(
        "SELECT ii.*, i.invoice_number, i.invoice_date, "
        "s.name AS supplier_name "
        "FROM invoice_items ii "
        "JOIN invoices i ON i.id = ii.invoice_id "
        "JOIN suppliers s ON s.id = i.supplier_id "
        "WHERE i.project_id = :pid"
    )
    with get_engine().connect() as conn:
        return _rows_to_dicts(conn.execute(sql, {"pid": project_id}).fetchall())
