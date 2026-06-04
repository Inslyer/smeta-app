"""Главный файл приложения — Streamlit UI.

Запуск (Mac):    bash start.sh
Запуск (Win):    start.bat
Напрямую:        streamlit run app.py

Версия БЕЗ AI-сопоставления — для проверки что Streamlit стабильно работает.
Полная версия с AI лежит в app_with_ai.py.
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
from core.models import Estimate

ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(ENV_PATH)
db.init_db()

st.set_page_config(
    page_title="Сводная смета",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Сводная смета — автоматизация закупки")


def _save_upload(uploaded_file) -> Path:
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
    st.caption(f"База цен: `{db.DEFAULT_DB_PATH.name}`")

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
tab_upload, tab_summary = st.tabs(["📤 Загрузка смет", "📊 Сводная смета"])

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
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric("Сумма клиента (с НДС)", f"{client.total:,.0f} ₽".replace(",", " "))
        if contractor and contractor.total:
            col2.metric(
                "Подрядчик (без НДС)",
                f"{contractor.total:,.0f} ₽".replace(",", " "),
            )
        col3.metric("Позиций у клиента", len(client.items))

        st.divider()
        st.info("AI-сопоставление будет включено на следующем шаге, после того как убедимся что базовое приложение стабильно работает.")
