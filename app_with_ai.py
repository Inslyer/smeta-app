"""Главный файл приложения — Streamlit UI.

Запуск (Mac):    bash start.sh
Запуск (Win):    start.bat
Напрямую:        streamlit run app.py
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from core import db
from core.parser_client import parse_client_estimate
from core.parser_contractor import parse_contractor_estimate
from core.exporter import build_workbook, filename_for
from core.matcher import match_estimates
from core.models import Estimate, Match, SupplierInvoice
from core.parser_supplier_invoice import parse_invoice

# Явный путь к .env — иначе Streamlit его не находит при запуске не из cwd проекта
ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(ENV_PATH)
db.init_db()

st.set_page_config(
    page_title="Сводная смета",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Сводная смета — автоматизация закупки")

if not os.getenv("ANTHROPIC_API_KEY"):
    st.error(
        "Не найден ANTHROPIC_API_KEY. Создайте файл `.env` рядом с `app.py` "
        "и пропишите туда ключ от Claude API. См. `.env.example`."
    )
    st.stop()


def _save_upload(uploaded_file) -> Path:
    """Сохранить uploaded_file во временный файл и вернуть путь."""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return Path(tmp.name)


def _items_to_df(estimate: Estimate) -> pd.DataFrame:
    rows = []
    for it in estimate.items:
        rows.append(
            {
                "Раздел": it.section,
                "№": it.number,
                "Тип": it.kind,
                "Наименование": it.name,
                "Ед.": it.unit,
                "Кол-во": it.quantity,
                "Цена мат.": it.price_material,
                "Цена раб.": it.price_work,
                "Сумма мат.": it.sum_material,
                "Сумма раб.": it.sum_work,
                "Итого": it.sum_total,
                "Компонентов": len(it.components) if it.components else 0,
            }
        )
    return pd.DataFrame(rows)


# -------------------------------------------------------------------- Sidebar
with st.sidebar:
    st.header("⚙️ Настройки")
    st.caption(f"Модель: `{os.getenv('CLAUDE_MODEL', 'claude-sonnet-4-6')}`")
    st.caption(f"База цен: `{db.db_label()}`")

    st.divider()
    st.subheader("Поставщики и подрядчики")
    suppliers = db.list_suppliers()
    if suppliers:
        for s in suppliers:
            policy = "🔒 фикс" if s["price_policy"] == "fixed" else "📈 переменные"
            st.write(f"• **{s['name']}** ({s['kind']}, {policy})")
    else:
        st.info("Пока никого. Добавьте ниже.")

    with st.expander("➕ Добавить"):
        with st.form("add_supplier", clear_on_submit=True):
            name = st.text_input("Название")
            kind = st.selectbox("Тип", ["material", "work"],
                                format_func=lambda x: "Материалы" if x == "material" else "Работы (СМР)")
            policy = st.selectbox("Политика цен", ["volatile", "fixed"],
                                  format_func=lambda x: "Переменные" if x == "volatile" else "Фиксированные")
            notes = st.text_area("Заметки", height=60)
            if st.form_submit_button("Сохранить"):
                if name.strip():
                    db.upsert_supplier(name.strip(), kind, policy, notes or None)
                    st.success(f"Добавлен: {name}")
                    st.rerun()


# --------------------------------------------------------------------- Main
tab_upload, tab_summary, tab_invoices = st.tabs([
    "📤 Загрузка смет",
    "📊 Сводная смета",
    "📦 Счета поставщиков",
])

with tab_upload:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Смета клиента")
        st.caption("Excel со сметой, которую вы выставляете заказчику.")
        client_file = st.file_uploader(
            "Перетащите файл сюда",
            type=["xlsx", "xls"],
            key="client_uploader",
        )

    with col2:
        st.subheader("Смета подрядчика (СМР)")
        st.caption("Excel с КП от подрядчика на монтажные работы.")
        contractor_file = st.file_uploader(
            "Перетащите файл сюда",
            type=["xlsx", "xls"],
            key="contractor_uploader",
        )

    st.divider()

    if client_file:
        try:
            client_path = _save_upload(client_file)
            client_estimate = parse_client_estimate(client_path)
            st.session_state["client_estimate"] = client_estimate
            st.success(
                f"✅ Смета клиента: **{len(client_estimate.items)}** позиций, "
                f"итого **{client_estimate.total:,.2f} ₽** с НДС".replace(",", " ")
            )
            with st.expander("Подробно — позиции сметы клиента"):
                st.dataframe(_items_to_df(client_estimate), width="stretch", hide_index=True)
        except Exception as e:
            st.error(f"Не удалось распарсить смету клиента: {e}")

    if contractor_file:
        try:
            contractor_path = _save_upload(contractor_file)
            contractor_estimate = parse_contractor_estimate(contractor_path)
            st.session_state["contractor_estimate"] = contractor_estimate
            total_str = (
                f"{contractor_estimate.total:,.2f} ₽".replace(",", " ")
                if contractor_estimate.total else "не указан"
            )
            st.success(
                f"✅ Смета подрядчика: **{len(contractor_estimate.items)}** позиций, "
                f"итого **{total_str}** без НДС"
            )
            with st.expander("Подробно — позиции сметы подрядчика"):
                st.dataframe(_items_to_df(contractor_estimate), width="stretch", hide_index=True)
        except Exception as e:
            st.error(f"Не удалось распарсить смету подрядчика: {e}")

with tab_summary:
    client = st.session_state.get("client_estimate")
    contractor = st.session_state.get("contractor_estimate")

    if not client:
        st.info("Загрузите смету клиента на вкладке «Загрузка смет».")
    elif not contractor:
        st.info("Загрузите смету подрядчика, чтобы запустить сопоставление работ.")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric(
            "Сумма клиента (с НДС)",
            f"{client.total:,.0f} ₽".replace(",", " "),
        )
        col2.metric(
            "Подрядчик (без НДС)",
            f"{contractor.total:,.0f} ₽".replace(",", " ") if contractor.total else "—",
        )
        col3.metric("Позиций клиента / подрядчика", f"{len(client.items)} / {len(contractor.items)}")

        st.divider()
        st.subheader("🔗 Сопоставление работ — AI")

        action_col, info_col = st.columns([1, 3])
        with action_col:
            if st.button("✨ Сопоставить позиции", type="primary"):
                with st.spinner("Claude сопоставляет позиции..."):
                    try:
                        matches, meta = match_estimates(client, contractor)
                        st.session_state["matches"] = matches
                        st.session_state["match_meta"] = meta
                    except Exception as e:
                        st.error(f"Ошибка сопоставления: {e}")
            if st.button("🗑 Очистить кэш"):
                st.session_state.pop("matches", None)
                st.session_state.pop("match_meta", None)

        with info_col:
            meta = st.session_state.get("match_meta")
            if meta:
                if meta.get("from_cache"):
                    st.caption("📦 Результат из кэша (та же пара смет уже была обработана)")
                else:
                    tin = meta.get("input_tokens", "?")
                    tout = meta.get("output_tokens", "?")
                    st.caption(
                        f"🧠 Модель: `{meta.get('model', '?')}` · "
                        f"токены: in={tin}, out={tout}"
                    )

        matches: list[Match] | None = st.session_state.get("matches")
        if matches:
            rows = []
            for m in matches:
                c_item = client.items[m.client_idx]
                p_item = (
                    contractor.items[m.contractor_idx]
                    if m.contractor_idx is not None else None
                )
                if m.confidence >= 0.9:
                    badge = "🟢"
                elif m.confidence >= 0.7:
                    badge = "🟡"
                elif m.confidence >= 0.4:
                    badge = "🟠"
                else:
                    badge = "🔴"
                if p_item is None:
                    badge = "⚪️"
                rows.append({
                    "Уверен.": f"{badge} {m.confidence:.0%}",
                    "Раздел клиента": c_item.section,
                    "Позиция клиента": c_item.name,
                    "Ед.": c_item.unit,
                    "Кол-во": c_item.quantity,
                    "Цена работы клиента": c_item.price_work,
                    "Соответствие у подрядчика":
                        p_item.name if p_item else "— нет —",
                    "Цена подрядчика": p_item.price_work if p_item else None,
                    "Обоснование": m.reason,
                })
            st.dataframe(
                pd.DataFrame(rows),
                width="stretch",
                hide_index=True,
            )

            green = sum(1 for m in matches if m.confidence >= 0.9 and m.contractor_idx is not None)
            yellow = sum(1 for m in matches if 0.7 <= m.confidence < 0.9 and m.contractor_idx is not None)
            orange = sum(1 for m in matches if 0.4 <= m.confidence < 0.7 and m.contractor_idx is not None)
            none = sum(1 for m in matches if m.contractor_idx is None)
            st.caption(
                f"🟢 уверенных: **{green}**  ·  🟡 с оговорками: **{yellow}**  ·  "
                f"🟠 спорных: **{orange}**  ·  ⚪️ без пары: **{none}**"
            )

            st.divider()
            st.subheader("📥 Скачать сводную смету")
            st.caption(
                "Excel-файл с 4 листами: «Сводная смета» (с маржой и светофором), "
                "«Аналитика», и исходники обеих смет для аудита."
            )
            try:
                xlsx_bytes = build_workbook(client, contractor, matches)
                st.download_button(
                    label="💾 Скачать Excel со сводной сметой",
                    data=xlsx_bytes,
                    file_name=filename_for(client),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as e:
                st.error(f"Не удалось собрать Excel: {e}")
        else:
            st.info("Нажмите «Сопоставить позиции», чтобы Claude построил пары.")


# =====================================================================
# Вкладка: Счета поставщиков
# =====================================================================
def _invoice_items_to_df(invoice: SupplierInvoice) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "№": it.line_no,
            "Артикул поставщика": it.article_supplier,
            "Артикул производителя": it.article_manufacturer,
            "Наименование": it.name,
            "Ед.": it.unit,
            "Кол-во": it.quantity,
            "Цена за ед.": it.unit_price,
            "НДС %": f"{it.vat_rate*100:.0f}%" if it.vat_rate else "—",
            "С НДС": "да" if it.vat_included else "нет",
            "Сумма": round(it.quantity * it.unit_price, 2),
        }
        for it in invoice.items
    ])


with tab_invoices:
    st.subheader("📦 Счета поставщиков по проектам")
    st.caption(
        "Закупщик загружает счёт от поставщика — Claude парсит позиции, "
        "вы подтверждаете, и цены сохраняются в базе для подстановки в сводную смету."
    )

    # --- Проекты ----------------------------------------------------------
    projects = db.list_projects()
    project_names = ["— выбрать проект —"] + [p["name"] for p in projects] + ["➕ Новый проект"]

    col_p1, col_p2 = st.columns([2, 3])
    with col_p1:
        chosen = st.selectbox("Проект", project_names, key="invoice_project_sel")
    with col_p2:
        new_project_name = ""
        if chosen == "➕ Новый проект":
            new_project_name = st.text_input(
                "Название нового проекта",
                placeholder="название объекта",
            )
            if new_project_name.strip() and st.button("Создать проект"):
                db.upsert_project(new_project_name.strip())
                st.success(f"Создан проект: {new_project_name}")
                st.rerun()

    project_id: int | None = None
    project_name: str | None = None
    if chosen not in ("— выбрать проект —", "➕ Новый проект"):
        project = db.get_project_by_name(chosen)
        if project:
            project_id = project["id"]
            project_name = project["name"]

    if not project_id:
        st.info("Выберите или создайте проект, чтобы загружать к нему счета.")
        st.stop()

    st.divider()

    # --- Загрузка нового счёта -------------------------------------------
    st.markdown("### ⬆️ Загрузить новый счёт")

    suppliers_material = [s for s in db.list_suppliers() if s["kind"] == "material"]
    if not suppliers_material:
        st.warning(
            "Нет ни одного поставщика материалов. Добавьте его в боковой "
            "панели (раздел «Поставщики и подрядчики»)."
        )
    else:
        col_s, col_f = st.columns([1, 2])
        with col_s:
            supplier_options = {s["name"]: s["id"] for s in suppliers_material}
            chosen_supplier = st.selectbox(
                "Поставщик",
                list(supplier_options.keys()),
                key="invoice_supplier_sel",
            )
            supplier_id = supplier_options[chosen_supplier]
        with col_f:
            invoice_file = st.file_uploader(
                "Файл счёта (PDF или Excel)",
                type=["pdf", "xlsx", "xls"],
                key="invoice_uploader",
            )

        if invoice_file is not None:
            cache_key = f"parsed_invoice::{invoice_file.name}::{invoice_file.size}"
            parsed: SupplierInvoice | None = st.session_state.get(cache_key)

            if parsed is None:
                with st.spinner("Claude читает счёт..."):
                    try:
                        tmp_path = _save_upload(invoice_file)
                        parsed = parse_invoice(tmp_path)
                        st.session_state[cache_key] = parsed
                    except Exception as e:
                        st.error(f"Не удалось распарсить счёт: {e}")
                        parsed = None

            if parsed is not None:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Поставщик", parsed.supplier_name[:24])
                c2.metric("Счёт", parsed.invoice_number or "—")
                c3.metric("Дата", parsed.invoice_date or "—")
                total_disp = (
                    f"{parsed.total_with_vat:,.0f} ₽".replace(",", " ")
                    if parsed.total_with_vat else "—"
                )
                c4.metric("Итого с НДС", total_disp)

                if parsed.project_tag:
                    st.caption(
                        f"ℹ️ В счёте указано примечание / тег: «{parsed.project_tag}» "
                        f"(носит справочный характер — привязка к проекту делается "
                        f"вашим выбором выше)."
                    )

                st.markdown(f"**Позиции ({len(parsed.items)}):**")
                st.dataframe(
                    _invoice_items_to_df(parsed),
                    width="stretch",
                    hide_index=True,
                )

                save_col, _ = st.columns([1, 4])
                with save_col:
                    if st.button("💾 Сохранить счёт", type="primary",
                                  key=f"save_{cache_key}"):
                        try:
                            invoice_id = db.save_invoice(
                                supplier_id=supplier_id,
                                project_id=project_id,
                                invoice_number=parsed.invoice_number,
                                invoice_date=parsed.invoice_date,
                                total_without_vat=parsed.total_without_vat,
                                total_with_vat=parsed.total_with_vat,
                                source_file=parsed.source_file,
                                items=[
                                    {
                                        "line_no": it.line_no,
                                        "name": it.name,
                                        "article_supplier": it.article_supplier,
                                        "article_manufacturer": it.article_manufacturer,
                                        "unit": it.unit,
                                        "quantity": it.quantity,
                                        "unit_price": it.unit_price,
                                        "vat_rate": it.vat_rate,
                                        "vat_included": it.vat_included,
                                    }
                                    for it in parsed.items
                                ],
                            )
                            st.success(
                                f"✅ Счёт сохранён в проект «{project_name}» "
                                f"(id={invoice_id})."
                            )
                            st.session_state.pop(cache_key, None)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Не удалось сохранить: {e}")

    st.divider()

    # --- Список ранее загруженных счетов ---------------------------------
    st.markdown(f"### 📋 Загруженные счета по проекту «{project_name}»")
    invoices = db.list_invoices(project_id=project_id)
    if not invoices:
        st.info("По этому проекту ещё нет загруженных счетов.")
    else:
        for inv in invoices:
            header = (
                f"**{inv['supplier_name']}** · "
                f"счёт №{inv['invoice_number'] or '—'} от "
                f"{inv['invoice_date'] or '—'} · "
                f"итого с НДС: "
                f"{(inv['total_with_vat'] or 0):,.0f} ₽".replace(",", " ")
            )
            with st.expander(header):
                items = db.list_invoice_items(inv["id"])
                if items:
                    df = pd.DataFrame([dict(r) for r in items])
                    show_cols = [
                        "line_no", "article_supplier", "article_manufacturer",
                        "name", "unit", "quantity", "unit_price",
                        "vat_rate", "vat_included",
                    ]
                    df = df[[c for c in show_cols if c in df.columns]]
                    df = df.rename(columns={
                        "line_no": "№",
                        "article_supplier": "Арт. поставщика",
                        "article_manufacturer": "Арт. производителя",
                        "name": "Наименование",
                        "unit": "Ед.",
                        "quantity": "Кол-во",
                        "unit_price": "Цена",
                        "vat_rate": "НДС",
                        "vat_included": "С НДС",
                    })
                    st.dataframe(df, width="stretch", hide_index=True)
                if st.button("🗑 Удалить счёт", key=f"del_inv_{inv['id']}"):
                    db.delete_invoice(inv["id"])
                    st.rerun()
