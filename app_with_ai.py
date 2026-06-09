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
from core.matcher import (
    classify_contractor_extras,
    match_estimates,
    match_materials,
)
from core.models import (
    ContractorExtra,
    Estimate,
    MaterialMatch,
    Match,
    SupplierInvoice,
)
from core.parser_supplier_invoice import parse_invoice

# Явный путь к .env — иначе Streamlit его не находит при запуске не из cwd проекта
ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(ENV_PATH)

# Мост Streamlit secrets → os.environ.
# На Streamlit Cloud переменные из Settings → Secrets доступны через st.secrets.
# Прокидываем их в окружение, чтобы код, читающий os.getenv(...), работал
# одинаково локально и в облаке.
_secrets_loaded: list[str] = []
_secrets_error: str | None = None
try:
    # st.secrets ведёт себя как mapping — итерируемся по ключам напрямую
    for _k in list(st.secrets):
        try:
            _v = st.secrets[_k]
            if isinstance(_v, (str, int, float, bool)):
                os.environ.setdefault(_k, str(_v))
                _secrets_loaded.append(_k)
        except Exception as _e:
            _secrets_error = f"{_k}: {_e}"
except Exception as _e:
    # st.secrets вообще недоступен (локальный запуск без secrets.toml)
    _secrets_error = f"st.secrets unavailable: {_e}"

db.init_db()

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_MARK = ASSETS_DIR / "graftio_symbol.svg"
LOGO_FULL = ASSETS_DIR / "graftio_planet.svg"

st.set_page_config(
    page_title="Генрих",
    page_icon=str(LOGO_MARK) if LOGO_MARK.exists() else "🪐",
    layout="wide",
)

# ----------- Фирменный CSS Графтио (тёмная тема) --------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], [data-testid="stMarkdownContainer"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* Чёрный фон всему приложению */
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
    background-color: #0E0E0E;
    color: #F2F2F2;
}

[data-testid="stHeader"] { background: transparent; }

/* Заголовки и текст — белые */
h1, h2, h3, h4, h5, h6,
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] * ,
.stMarkdown, .stMarkdown p, .stMarkdown li, .stMarkdown span {
    color: #F2F2F2;
}
[data-testid="stCaptionContainer"], small, .stCaption {
    color: #A0A0A0 !important;
}

/* Primary кнопки — лайм-зелёный градиент с чёрным текстом */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #7BBE2F 0%, #A8DC1A 50%, #F4E821 100%);
    color: #0E0E0E;
    border: none;
    font-weight: 700;
    box-shadow: 0 2px 8px rgba(168,220,26,0.25);
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #6FA82A 0%, #9CCC18 50%, #E8DD20 100%);
    color: #0E0E0E;
    box-shadow: 0 3px 12px rgba(168,220,26,0.4);
}

/* Обычные кнопки — тёмные с тонкой обводкой */
.stButton > button:not([kind="primary"]) {
    background: #1A1A1A;
    color: #F2F2F2;
    border: 1px solid #333;
}
.stButton > button:not([kind="primary"]):hover {
    background: #232323;
    border-color: #A8DC1A;
    color: #F2F2F2;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #131313 !important;
    border-right: 1px solid #232323;
}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {
    color: #F2F2F2;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 1px solid #232323;
}
.stTabs [data-baseweb="tab"] {
    padding: 10px 20px;
    border-radius: 8px 8px 0 0;
    color: #A0A0A0;
    background: transparent;
}
.stTabs [aria-selected="true"] {
    background: #1A1A1A;
    color: #F2F2F2 !important;
    border-bottom: 3px solid #A8DC1A !important;
}

/* Метрики */
[data-testid="stMetric"] {
    background: #1A1A1A;
    padding: 14px 18px;
    border-radius: 10px;
    border-left: 3px solid #A8DC1A;
}
[data-testid="stMetricLabel"] { color: #A0A0A0 !important; }
[data-testid="stMetricValue"] { color: #F2F2F2 !important; }

/* Поля ввода и selectbox */
.stTextInput input, .stTextArea textarea, .stNumberInput input,
[data-baseweb="select"] > div {
    background-color: #1A1A1A !important;
    color: #F2F2F2 !important;
    border-color: #333 !important;
}

/* DataFrame */
[data-testid="stDataFrame"] {
    background: #1A1A1A;
    border-radius: 8px;
}

/* Expanders */
[data-testid="stExpander"] {
    background: #1A1A1A;
    border: 1px solid #232323;
    border-radius: 8px;
}
[data-testid="stExpander"] summary { color: #F2F2F2; }

/* File uploader */
[data-testid="stFileUploader"] section {
    background: #1A1A1A;
    border: 2px dashed #333;
    border-radius: 10px;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #A8DC1A;
}

/* Alerts (info/success/warning/error) */
[data-testid="stAlert"] { border-radius: 8px; }

/* Divider */
hr { border-color: #232323; }

/* Шапка страницы — место для логотипа */
.graftio-header {
    display: flex;
    align-items: center;
    gap: 18px;
    margin-bottom: 6px;
    padding-bottom: 12px;
    border-bottom: 2px solid #A8DC1A;
}
.graftio-header .mark { width: 56px; height: 56px; flex-shrink: 0; }
.graftio-header h1 {
    margin: 0 !important;
    padding: 0 !important;
    color: #F2F2F2 !important;
    font-weight: 800;
    letter-spacing: -0.5px;
    font-size: 2.4rem;
}
.graftio-tagline {
    color: #A0A0A0;
    font-size: 0.9em;
    margin-top: 6px;
    margin-bottom: 24px;
    font-weight: 500;
}
</style>
""",
    unsafe_allow_html=True,
)

# --- Заголовок с логотипом ---
if LOGO_MARK.exists():
    logo_svg = LOGO_MARK.read_text(encoding="utf-8")
    st.markdown(
        f"""
<div class="graftio-header">
    <div class="mark">{logo_svg}</div>
    <h1>Генрих</h1>
</div>
<div class="graftio-tagline">Сводная смета — автоматизация закупки</div>
""",
        unsafe_allow_html=True,
    )
else:
    st.title("Генрих")

if not os.getenv("ANTHROPIC_API_KEY"):
    st.error(
        "Не найден ANTHROPIC_API_KEY. Создайте файл `.env` рядом с `app.py` "
        "и пропишите туда ключ Anthropic API (его использует Генрих). "
        "См. `.env.example`."
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
    if LOGO_MARK.exists():
        logo_svg_sidebar = LOGO_MARK.read_text(encoding="utf-8")
        st.markdown(
            f"""
<div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
    <div style="width:40px;height:40px;">{logo_svg_sidebar}</div>
    <div style="font-weight:800; font-size:1.3em; color:#F2F2F2; letter-spacing:-0.5px;">Генрих</div>
</div>
""",
            unsafe_allow_html=True,
        )
        st.divider()

    st.header("⚙️ Настройки")
    st.caption(f"Модель: `{os.getenv('CLAUDE_MODEL', 'claude-sonnet-4-6')}`")
    st.caption(f"База цен: `{db.db_label()}`")

    with st.expander("🩺 Диагностика окружения"):
        st.write({
            "ANTHROPIC_API_KEY": "✓ задан" if os.getenv("ANTHROPIC_API_KEY") else "✗ нет",
            "DATABASE_URL": "✓ задан" if os.getenv("DATABASE_URL") else "✗ нет (fallback на SQLite)",
            "secrets из st.secrets": _secrets_loaded or "пусто",
            "ошибка чтения secrets": _secrets_error or "—",
            "backend": db.db_label(),
        })

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
        st.subheader("Смета Заказчика")
        st.caption("Excel со сметой, которую вы выставляете Заказчику.")
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
            # Подменяем технические tmp-имена на оригинальное имя файла,
            # чтобы в Excel в колонке «Исполнитель» был человеко-читаемый источник.
            client_estimate.title = Path(client_file.name).stem
            client_estimate.file_name = client_file.name
            st.session_state["client_estimate"] = client_estimate
            st.session_state["client_xlsx_bytes"] = bytes(client_file.getbuffer())
            st.session_state["client_xlsx_name"] = client_file.name
            st.success(
                f"✅ Смета Заказчика: **{len(client_estimate.items)}** позиций, "
                f"итого **{client_estimate.total:,.2f} ₽** с НДС".replace(",", " ")
            )
            with st.expander("Подробно — позиции сметы Заказчика"):
                st.dataframe(_items_to_df(client_estimate), width="stretch", hide_index=True)
        except Exception as e:
            st.error(f"Не удалось распарсить смету Заказчика: {e}")

    if contractor_file:
        try:
            contractor_path = _save_upload(contractor_file)
            contractor_estimate = parse_contractor_estimate(contractor_path)
            contractor_estimate.title = Path(contractor_file.name).stem
            contractor_estimate.file_name = contractor_file.name
            st.session_state["contractor_estimate"] = contractor_estimate
            st.session_state["contractor_xlsx_bytes"] = bytes(contractor_file.getbuffer())
            st.session_state["contractor_xlsx_name"] = contractor_file.name
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
        st.info("Загрузите смету Заказчика на вкладке «Загрузка смет».")
    elif not contractor:
        st.info("Загрузите смету подрядчика, чтобы запустить сопоставление работ.")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric(
            "Сумма Заказчика (с НДС)",
            f"{client.total:,.0f} ₽".replace(",", " "),
        )
        col2.metric(
            "Подрядчик (без НДС)",
            f"{contractor.total:,.0f} ₽".replace(",", " ") if contractor.total else "—",
        )
        col3.metric("Позиций Заказчика / подрядчика", f"{len(client.items)} / {len(contractor.items)}")

        st.divider()
        st.subheader("🔗 Сопоставление работ — AI")

        action_col, info_col = st.columns([1, 3])
        with action_col:
            if st.button("✨ Сопоставить позиции", type="primary"):
                with st.spinner("Генрих сопоставляет позиции..."):
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
                    "Раздел Заказчика": c_item.section,
                    "Позиция Заказчика": c_item.name,
                    "Ед.": c_item.unit,
                    "Кол-во": c_item.quantity,
                    "Цена работы Заказчика": c_item.price_work,
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

            # ============================================================
            # Доп. позиции подрядчика — что осталось без 1-в-1 матча
            # ============================================================
            used_contractor_idxs = {m.contractor_idx for m in matches
                                      if m.contractor_idx is not None}
            unmatched_contractor = [
                i for i in range(len(contractor.items))
                if i not in used_contractor_idxs
            ]

            if unmatched_contractor:
                st.divider()
                st.subheader(
                    f"🔍 Остаток подрядчика — {len(unmatched_contractor)} поз. "
                    "без прямого матча"
                )
                st.caption(
                    "Это позиции подрядчика, которым не нашлось пары в смете "
                    "Заказчика. Генрих разберёт: какие из них — компоненты "
                    "крупных работ Заказчика, а какие — самостоятельные доп. "
                    "работы вне сметы."
                )

                col_e1, col_e2 = st.columns([1, 3])
                with col_e1:
                    if st.button("🧪 Классифицировать остаток",
                                  type="primary", key="run_extras"):
                        with st.spinner("Генрих разбирает остаток..."):
                            try:
                                extras_res, extras_meta = classify_contractor_extras(
                                    client, contractor, matches,
                                )
                                st.session_state["contractor_extras"] = extras_res
                                st.session_state["extras_meta"] = extras_meta
                            except Exception as e:
                                st.error(f"Ошибка: {e}")
                    if st.button("🗑 Сбросить", key="clear_extras"):
                        st.session_state.pop("contractor_extras", None)
                        st.session_state.pop("extras_meta", None)
                with col_e2:
                    em = st.session_state.get("extras_meta")
                    if em:
                        if em.get("from_cache"):
                            st.caption("📦 Из кэша")
                        else:
                            st.caption(
                                f"🧠 `{em.get('model','?')}` · "
                                f"токены in={em.get('input_tokens')}, "
                                f"out={em.get('output_tokens')}"
                            )

                extras: list[ContractorExtra] | None = \
                    st.session_state.get("contractor_extras")
                if extras:
                    rows = []
                    for e in extras:
                        c_p_item = contractor.items[e.contractor_idx]
                        parent_name = (
                            client.items[e.parent_client_idx].name
                            if e.parent_client_idx is not None
                            and 0 <= e.parent_client_idx < len(client.items)
                            else "—"
                        )
                        if e.kind == "included":
                            kind_badge = "🧩 компонент"
                        else:
                            kind_badge = "🟪 вне сметы"
                        rows.append({
                            "Тип": kind_badge,
                            "Уверен.": f"{e.confidence:.0%}",
                            "Позиция подрядчика": c_p_item.name,
                            "Кол-во": f"{c_p_item.quantity} {c_p_item.unit}",
                            "Цена": c_p_item.price_work,
                            "Сумма": (c_p_item.price_work or 0)
                                       * (c_p_item.quantity or 0),
                            "Входит в позицию Заказчика": parent_name,
                            "Обоснование": e.reason,
                        })
                    st.dataframe(
                        pd.DataFrame(rows),
                        width="stretch",
                        hide_index=True,
                    )

                    inc = sum(1 for e in extras if e.kind == "included")
                    ext = sum(1 for e in extras if e.kind == "extra")
                    inc_sum = sum(
                        (contractor.items[e.contractor_idx].price_work or 0)
                        * (contractor.items[e.contractor_idx].quantity or 0)
                        for e in extras if e.kind == "included"
                    )
                    ext_sum = sum(
                        (contractor.items[e.contractor_idx].price_work or 0)
                        * (contractor.items[e.contractor_idx].quantity or 0)
                        for e in extras if e.kind == "extra"
                    )
                    st.caption(
                        f"🧩 компонентов крупных работ: **{inc}** "
                        f"({inc_sum:,.0f} ₽)  ·  ".replace(",", " ")
                        + f"🟪 вне сметы Заказчика: **{ext}** "
                        f"({ext_sum:,.0f} ₽)".replace(",", " ")
                    )

            # ============================================================
            # Сопоставление МАТЕРИАЛОВ клиента со счетами поставщиков
            # ============================================================
            st.divider()
            st.subheader("🧾 Цены закупки материалов — из счетов поставщиков")
            st.caption(
                "Выберите проект — Генрих сопоставит материалы из сметы Заказчика "
                "с позициями загруженных счетов и подставит цены закупки."
            )

            mat_projects = db.list_projects()
            if not mat_projects:
                st.info(
                    "Нет ни одного проекта. Создайте его на вкладке "
                    "«📦 Счета поставщиков» и загрузите туда счета."
                )
            else:
                mat_proj_options = ["— выбрать проект —"] + [
                    p["name"] for p in mat_projects
                ]
                mat_chosen = st.selectbox(
                    "Проект для подстановки цен закупки",
                    mat_proj_options,
                    key="materials_project_sel",
                )
                mat_project_id = None
                if mat_chosen != "— выбрать проект —":
                    mat_project = db.get_project_by_name(mat_chosen)
                    if mat_project:
                        mat_project_id = mat_project["id"]

                if mat_project_id:
                    supplier_rows = db.all_invoice_items_for_project(mat_project_id)
                    if not supplier_rows:
                        st.info(
                            f"По проекту «{mat_chosen}» ещё нет загруженных "
                            "счетов. Загрузите их на вкладке «📦 Счета "
                            "поставщиков»."
                        )
                    else:
                        st.caption(
                            f"Позиций в счетах проекта: **{len(supplier_rows)}**"
                        )
                        col_mat_act, col_mat_info = st.columns([1, 3])
                        with col_mat_act:
                            if st.button("🧮 Подобрать цены закупки",
                                          type="primary",
                                          key="run_match_materials"):
                                with st.spinner("Генрих сопоставляет материалы..."):
                                    try:
                                        m_matches, m_meta = match_materials(
                                            client, supplier_rows,
                                        )
                                        st.session_state["mat_matches"] = m_matches
                                        st.session_state["mat_meta"] = m_meta
                                        st.session_state["mat_supplier_rows"] = supplier_rows
                                    except Exception as e:
                                        st.error(f"Ошибка сопоставления: {e}")
                            if st.button("🗑 Сбросить", key="clear_mat_matches"):
                                for k in ("mat_matches", "mat_meta", "mat_supplier_rows"):
                                    st.session_state.pop(k, None)
                        with col_mat_info:
                            mm = st.session_state.get("mat_meta")
                            if mm:
                                if mm.get("from_cache"):
                                    st.caption("📦 Из кэша")
                                else:
                                    st.caption(
                                        f"🧠 Модель: `{mm.get('model', '?')}` · "
                                        f"токены: in={mm.get('input_tokens')}, "
                                        f"out={mm.get('output_tokens')}"
                                    )

                    m_matches: list[MaterialMatch] | None = st.session_state.get("mat_matches")
                    cached_supplier_rows = st.session_state.get("mat_supplier_rows", [])
                    if m_matches and cached_supplier_rows:
                        # Индексируем supplier_rows по id для быстрого lookup
                        by_id = {r["id"]: r for r in cached_supplier_rows}
                        mat_rows = []
                        for mm in m_matches:
                            c_item = client.items[mm.client_idx]
                            s_row = by_id.get(mm.invoice_item_id) if mm.invoice_item_id else None
                            if mm.confidence >= 0.9 and s_row is not None:
                                badge = "🟢"
                            elif mm.confidence >= 0.7 and s_row is not None:
                                badge = "🟡"
                            elif mm.confidence >= 0.4 and s_row is not None:
                                badge = "🟠"
                            else:
                                badge = "🔴" if s_row else "⚪️"
                            kind_badge = {"article": "🔑", "semantic": "💬",
                                           "none": "—"}.get(mm.match_kind, "")
                            mat_rows.append({
                                "Уверен.": f"{badge} {mm.confidence:.0%}",
                                "Тип": kind_badge,
                                "Раздел Заказчика": c_item.section,
                                "Материал Заказчика": c_item.name,
                                "Кол-во Заказчика": f"{c_item.quantity} {c_item.unit}",
                                "Цена Заказчика (мат.)": c_item.price_material,
                                "Позиция в счёте": s_row["name"] if s_row else "— нет —",
                                "Артикул": (s_row.get("article_supplier")
                                              or s_row.get("article_manufacturer"))
                                              if s_row else None,
                                "Цена закупки": s_row.get("unit_price") if s_row else None,
                                "Поставщик": s_row.get("supplier_name") if s_row else None,
                                "Счёт": (
                                    f"№{s_row.get('invoice_number')} от "
                                    f"{s_row.get('invoice_date')}"
                                ) if s_row else None,
                                "Обоснование": mm.reason,
                            })
                        st.dataframe(
                            pd.DataFrame(mat_rows),
                            width="stretch",
                            hide_index=True,
                        )
                        m_green = sum(1 for x in m_matches
                                       if x.confidence >= 0.9 and x.invoice_item_id)
                        m_article = sum(1 for x in m_matches
                                         if x.match_kind == "article")
                        m_none = sum(1 for x in m_matches if not x.invoice_item_id)
                        st.caption(
                            f"🔑 по артикулу: **{m_article}**  ·  🟢 уверенных: "
                            f"**{m_green}**  ·  ⚪️ без пары: **{m_none}**"
                        )

            st.divider()
            st.subheader("📥 Скачать сводную смету")
            st.caption(
                "Excel-файл с 4 листами: «Сводная смета» (с маржой и светофором), "
                "«Аналитика», и исходники обеих смет для аудита."
            )
            # Собираем список счетов для отдельных вкладок (если выбран проект
            # на этой же вкладке сводной сметы)
            invoices_with_items: list[tuple[dict, list[dict]]] = []
            mat_proj_id = None
            mat_proj_name = st.session_state.get("materials_project_sel")
            if mat_proj_name and mat_proj_name != "— выбрать проект —":
                _proj = db.get_project_by_name(mat_proj_name)
                if _proj:
                    mat_proj_id = _proj["id"]
            if mat_proj_id:
                for inv in db.list_invoices(project_id=mat_proj_id):
                    inv_items = db.list_invoice_items(inv["id"])
                    invoices_with_items.append((dict(inv), inv_items))

            try:
                xlsx_bytes = build_workbook(
                    client, contractor, matches,
                    material_matches=st.session_state.get("mat_matches"),
                    supplier_rows=st.session_state.get("mat_supplier_rows"),
                    client_xlsx_bytes=st.session_state.get("client_xlsx_bytes"),
                    contractor_xlsx_bytes=st.session_state.get("contractor_xlsx_bytes"),
                    invoices_with_items=invoices_with_items or None,
                    contractor_extras=st.session_state.get("contractor_extras"),
                )
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
            st.info("Нажмите «Сопоставить позиции», чтобы Генрих построил пары.")


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
        "Закупщик загружает счёт от поставщика — Генрих парсит позиции, "
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
                "Поставщик (для всех загружаемых счетов)",
                list(supplier_options.keys()),
                key="invoice_supplier_sel",
            )
            supplier_id = supplier_options[chosen_supplier]
        with col_f:
            invoice_files = st.file_uploader(
                "Файлы счетов (можно несколько — PDF или Excel)",
                type=["pdf", "xlsx", "xls"],
                key="invoice_uploader",
                accept_multiple_files=True,
            )

        if invoice_files:
            st.caption(
                f"Файлов в очереди: **{len(invoice_files)}**. "
                "Генрих читает каждый по очереди и кэширует результат."
            )

            parsed_list: list[tuple[str, SupplierInvoice | None, str | None]] = []
            for f in invoice_files:
                cache_key = f"parsed_invoice::{f.name}::{f.size}"
                parsed_one: SupplierInvoice | None = st.session_state.get(cache_key)
                err = None
                if parsed_one is None:
                    with st.spinner(f"Генрих читает «{f.name}»..."):
                        try:
                            tmp_path = _save_upload(f)
                            parsed_one = parse_invoice(tmp_path)
                            st.session_state[cache_key] = parsed_one
                        except Exception as e:
                            err = str(e)
                            parsed_one = None
                parsed_list.append((cache_key, parsed_one, err))

            # Превью всех распарсенных
            for cache_key, parsed_one, err in parsed_list:
                f_name = cache_key.split("::")[1]
                if err:
                    st.error(f"❌ {f_name}: {err}")
                    continue
                if parsed_one is None:
                    continue
                header = (
                    f"📄 **{f_name}** · {parsed_one.supplier_name[:24]} · "
                    f"счёт №{parsed_one.invoice_number or '—'} от "
                    f"{parsed_one.invoice_date or '—'} · "
                    f"{len(parsed_one.items)} поз."
                )
                with st.expander(header):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Поставщик", parsed_one.supplier_name[:24])
                    c2.metric("Счёт", parsed_one.invoice_number or "—")
                    c3.metric("Дата", parsed_one.invoice_date or "—")
                    total_disp = (
                        f"{parsed_one.total_with_vat:,.0f} ₽".replace(",", " ")
                        if parsed_one.total_with_vat else "—"
                    )
                    c4.metric("Итого с НДС", total_disp)
                    if parsed_one.project_tag:
                        st.caption(
                            f"ℹ️ В счёте указано примечание / тег: "
                            f"«{parsed_one.project_tag}» (справочно)."
                        )
                    st.dataframe(
                        _invoice_items_to_df(parsed_one),
                        width="stretch",
                        hide_index=True,
                    )

            ok_parsed = [(k, p) for k, p, e in parsed_list if p is not None and not e]
            if ok_parsed:
                col_save, col_clear, _ = st.columns([2, 1, 3])
                with col_save:
                    if st.button(
                        f"💾 Сохранить все ({len(ok_parsed)}) в проект "
                        f"«{project_name}»",
                        type="primary",
                        key="save_all_invoices",
                    ):
                        saved, errors = 0, []
                        for cache_key, parsed_one in ok_parsed:
                            try:
                                db.save_invoice(
                                    supplier_id=supplier_id,
                                    project_id=project_id,
                                    invoice_number=parsed_one.invoice_number,
                                    invoice_date=parsed_one.invoice_date,
                                    total_without_vat=parsed_one.total_without_vat,
                                    total_with_vat=parsed_one.total_with_vat,
                                    source_file=parsed_one.source_file,
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
                                        for it in parsed_one.items
                                    ],
                                )
                                saved += 1
                                st.session_state.pop(cache_key, None)
                            except Exception as e:
                                errors.append(
                                    f"{cache_key.split('::')[1]}: {e}"
                                )
                        if saved:
                            st.success(
                                f"✅ Сохранено счетов: **{saved}** в проект "
                                f"«{project_name}»."
                            )
                        for er in errors:
                            st.error(f"❌ {er}")
                        if saved and not errors:
                            st.rerun()
                with col_clear:
                    if st.button("🗑 Очистить очередь",
                                   key="clear_invoices_queue"):
                        for cache_key, _ in ok_parsed:
                            st.session_state.pop(cache_key, None)
                        st.rerun()

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
