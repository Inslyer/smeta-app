"""Генератор итогового Excel со сводной сметой и аналитикой.

Структура книги:
    1. «Сводная смета» — каждая позиция клиента разворачивается на ДВЕ строки
        (материал и работа отдельно), цены формулами, итоги SUM, светофор маржи.
    2. «Аналитика» — две таблицы: маржа по работам и по материалам по категориям.
    3. «Исходник клиента» / «Исходник подрядчика» — оригинальные сметы с итогами.

НДС РФ 2026 = 22%. Маржа считается без НДС.
Категории материалов: Лотки, Кабель, Зарядные станции, Шкафы, Другое.
"""
from __future__ import annotations

import io
import re
from datetime import date
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import Estimate, EstimateItem, Match


VAT_RATE = 0.22  # НДС РФ с 2026
VAT_DIVISOR = f"{1 + VAT_RATE:.2f}"  # для формул

# Палитра светофора по марже %.
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
RED = PatternFill("solid", fgColor="FFC7CE")
GREY = PatternFill("solid", fgColor="EEECEC")
SECTION_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="B4C7E7")
PROJECT_TOTAL_FILL = PatternFill("solid", fgColor="2E75B6")
MATERIAL_TINT = PatternFill("solid", fgColor="FFF7E6")  # лёгкий фон для строк материала
WORK_TINT = PatternFill("solid", fgColor="EAF7EA")      # лёгкий фон для строк работы

BORDER_THIN = Border(
    left=Side(style="thin", color="C0C0C0"),
    right=Side(style="thin", color="C0C0C0"),
    top=Side(style="thin", color="C0C0C0"),
    bottom=Side(style="thin", color="C0C0C0"),
)
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
PROJECT_FONT = Font(bold=True, color="FFFFFF")


# ---------------------------------------------------------- Классификация материалов
CATEGORY_RULES = [
    ("Лотки", [r"лоток", r"лотка", r"разделитель лотк", r"крышка лотк",
               r"угол лотк", r"заглушк.*лотк", r"кабельнес"]),
    ("Кабель и проводка", [r"кабель", r"провод", r"витая пара",
                            r"коннектор", r"наконечник", r"муфт"]),
    ("Зарядные станции", [r"эзс", r"зарядн", r"пур[\-\s]?эзс", r"щсу[\-\s]?эзс"]),
    ("Шкафы и щиты", [r"шкаф", r"щит", r"сдупэм", r"щсу", r"корпус"]),
]


def classify_material(name: str) -> str:
    n = name.lower()
    for category, patterns in CATEGORY_RULES:
        for pat in patterns:
            if re.search(pat, n):
                return category
    return "Другое"


def _margin_fill(margin_pct: Optional[float]) -> PatternFill:
    if margin_pct is None:
        return GREY
    if margin_pct >= 0.25:
        return GREEN
    if margin_pct >= 0.10:
        return YELLOW
    return RED


# ---------------------------------------------------------- Сводная смета
# Колонки листа 1
COLUMNS = [
    ("Раздел", 28),
    ("№", 5),
    ("Тип строки", 11),
    ("Наименование", 55),
    ("Ед.", 8),
    ("Кол-во", 10),
    ("Цена клиента (с НДС)", 16),
    ("Сумма клиента (с НДС)", 18),
    ("Сумма клиента (без НДС)", 18),
    ("Исполнитель / Поставщик", 24),
    ("Цена закупки (без НДС)", 18),
    ("Сумма закупки (без НДС)", 18),
    ("Маржа, руб.", 14),
    ("Маржа, %", 11),
    ("Уверенность AI", 14),
    ("Комментарий AI", 50),
]
COL = {title: idx + 1 for idx, (title, _) in enumerate(COLUMNS)}

C_QTY = COL["Кол-во"]                              # F
C_UNIT_PRICE_CLIENT = COL["Цена клиента (с НДС)"]  # G
C_SUM_CLIENT_GROSS = COL["Сумма клиента (с НДС)"]  # H
C_SUM_CLIENT_NET = COL["Сумма клиента (без НДС)"]  # I
C_UNIT_PRICE_CONTR = COL["Цена закупки (без НДС)"]  # K
C_SUM_CONTR_NET = COL["Сумма закупки (без НДС)"]   # L
C_MARGIN_ABS = COL["Маржа, руб."]                  # M
C_MARGIN_PCT = COL["Маржа, %"]                     # N


def _col_letter(col_idx: int) -> str:
    return get_column_letter(col_idx)


def _write_headers(ws: Worksheet, columns=COLUMNS):
    for col_idx, (title, width) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 38
    ws.freeze_panes = "D2"


def _apply_border(ws, row_idx: int, ncols: int):
    for c in range(1, ncols + 1):
        ws.cell(row=row_idx, column=c).border = BORDER_THIN


def _section_row(ws, row_idx: int, section: str, ncols: int):
    for c in range(1, ncols + 1):
        ws.cell(row=row_idx, column=c).fill = SECTION_FILL
        ws.cell(row=row_idx, column=c).border = BORDER_THIN
    ws.cell(row=row_idx, column=1, value=section).font = Font(bold=True)
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=ncols)


def _money_format(cell):
    cell.number_format = "#,##0.00 ₽"


def _pct_format(cell):
    cell.number_format = "0.0%"


def _build_summary_sheet(ws: Worksheet, client: Estimate, contractor: Estimate,
                          matches: list[Match]):
    ws.title = "Сводная смета"
    _write_headers(ws)
    ncols = len(COLUMNS)

    match_by_client = {m.client_idx: m for m in matches}
    sections_in_order: list[str] = []
    seen = set()
    for it in client.items:
        if it.section not in seen:
            sections_in_order.append(it.section)
            seen.add(it.section)

    row = 2
    section_total_rows: list[tuple[str, int, list[int]]] = []  # (раздел, row_total, item_rows)

    for section in sections_in_order:
        _section_row(ws, row, section, ncols)
        row += 1

        section_item_rows: list[int] = []

        for c_idx, item in enumerate(client.items):
            if item.section != section:
                continue

            match = match_by_client.get(c_idx)
            contractor_item = (
                contractor.items[match.contractor_idx]
                if match and match.contractor_idx is not None else None
            )

            qty = item.quantity or 0
            # Решаем — какие строки выводить
            has_material = item.price_material is not None or item.sum_material is not None
            has_work = item.price_work is not None or item.sum_work is not None
            kinds_to_emit: list[str] = []
            if has_material:
                kinds_to_emit.append("material")
            if has_work:
                kinds_to_emit.append("work")
            if not kinds_to_emit:
                kinds_to_emit.append(item.kind if item.kind != "composite" else "material")

            for kind in kinds_to_emit:
                row_letter_qty = f"{_col_letter(C_QTY)}{row}"
                row_letter_unit_client = f"{_col_letter(C_UNIT_PRICE_CLIENT)}{row}"
                row_letter_sum_client_gross = f"{_col_letter(C_SUM_CLIENT_GROSS)}{row}"
                row_letter_sum_client_net = f"{_col_letter(C_SUM_CLIENT_NET)}{row}"
                row_letter_unit_contr = f"{_col_letter(C_UNIT_PRICE_CONTR)}{row}"
                row_letter_sum_contr = f"{_col_letter(C_SUM_CONTR_NET)}{row}"
                row_letter_margin = f"{_col_letter(C_MARGIN_ABS)}{row}"

                ws.cell(row=row, column=1, value=item.section)
                ws.cell(row=row, column=2, value=item.number)
                ws.cell(row=row, column=3,
                        value="Материал" if kind == "material" else "Работа")
                ws.cell(row=row, column=4, value=item.name).alignment = \
                    Alignment(wrap_text=True, vertical="top")
                ws.cell(row=row, column=5, value=item.unit)
                ws.cell(row=row, column=6, value=qty).number_format = "#,##0.00"

                unit_price = (item.price_material if kind == "material"
                              else item.price_work)
                ws.cell(row=row, column=C_UNIT_PRICE_CLIENT, value=unit_price)
                _money_format(ws.cell(row=row, column=C_UNIT_PRICE_CLIENT))

                # Сумма клиента с НДС — формула qty * unit
                if unit_price is not None:
                    ws.cell(
                        row=row, column=C_SUM_CLIENT_GROSS,
                        value=f"={row_letter_qty}*{row_letter_unit_client}",
                    )
                _money_format(ws.cell(row=row, column=C_SUM_CLIENT_GROSS))

                # Сумма без НДС — формула с НДС / (1+ставка)
                ws.cell(
                    row=row, column=C_SUM_CLIENT_NET,
                    value=f"={row_letter_sum_client_gross}/{VAT_DIVISOR}",
                )
                _money_format(ws.cell(row=row, column=C_SUM_CLIENT_NET))

                # Закупка — только для работы и если есть матч
                if kind == "work" and contractor_item:
                    ws.cell(row=row, column=10,
                            value=f"Подрядчик: {contractor.title}")
                    ws.cell(row=row, column=C_UNIT_PRICE_CONTR,
                            value=contractor_item.price_work)
                    _money_format(ws.cell(row=row, column=C_UNIT_PRICE_CONTR))
                    ws.cell(
                        row=row, column=C_SUM_CONTR_NET,
                        value=f"={row_letter_qty}*{row_letter_unit_contr}",
                    )
                    _money_format(ws.cell(row=row, column=C_SUM_CONTR_NET))

                    ws.cell(
                        row=row, column=C_MARGIN_ABS,
                        value=f"={row_letter_sum_client_net}-{row_letter_sum_contr}",
                    )
                    _money_format(ws.cell(row=row, column=C_MARGIN_ABS))

                    margin_pct_cell = ws.cell(
                        row=row, column=C_MARGIN_PCT,
                        value=(
                            f"=IF({row_letter_sum_client_net}=0,0,"
                            f"{row_letter_margin}/{row_letter_sum_client_net})"
                        ),
                    )
                    _pct_format(margin_pct_cell)

                    # Светофор — по формуле не покрасишь, посчитаю Python-стороной
                    if item.sum_work and contractor_item.price_work:
                        client_net = (item.sum_work or 0) / (1 + VAT_RATE)
                        c_net = (contractor_item.price_work or 0) * qty
                        margin_pct_value = ((client_net - c_net) / client_net
                                            if client_net else None)
                        margin_pct_cell.fill = _margin_fill(margin_pct_value)

                    if match:
                        conf_cell = ws.cell(row=row, column=15, value=match.confidence)
                        _pct_format(conf_cell)
                        ws.cell(row=row, column=16, value=match.reason).alignment = \
                            Alignment(wrap_text=True, vertical="top")
                else:
                    # Материал или работа без матча
                    ws.cell(
                        row=row, column=10,
                        value=("— нет поставщика —" if kind == "material" else "— не закрыто —"),
                    ).fill = GREY
                    ws.cell(row=row, column=C_MARGIN_PCT).fill = GREY

                # лёгкая подсветка типа строки
                ws.cell(row=row, column=3).fill = (
                    MATERIAL_TINT if kind == "material" else WORK_TINT
                )

                _apply_border(ws, row, ncols)
                section_item_rows.append(row)
                row += 1

        # Строка итога раздела
        if section_item_rows:
            first = section_item_rows[0]
            last = section_item_rows[-1]
            for c in range(1, ncols + 1):
                ws.cell(row=row, column=c).fill = TOTAL_FILL
                ws.cell(row=row, column=c).font = Font(bold=True)
                ws.cell(row=row, column=c).border = BORDER_THIN
            ws.cell(row=row, column=1, value=f"Итого по разделу: {section}")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            ws.cell(row=row, column=C_SUM_CLIENT_GROSS,
                    value=f"=SUM({_col_letter(C_SUM_CLIENT_GROSS)}{first}:"
                          f"{_col_letter(C_SUM_CLIENT_GROSS)}{last})")
            _money_format(ws.cell(row=row, column=C_SUM_CLIENT_GROSS))
            ws.cell(row=row, column=C_SUM_CLIENT_NET,
                    value=f"=SUM({_col_letter(C_SUM_CLIENT_NET)}{first}:"
                          f"{_col_letter(C_SUM_CLIENT_NET)}{last})")
            _money_format(ws.cell(row=row, column=C_SUM_CLIENT_NET))
            ws.cell(row=row, column=C_SUM_CONTR_NET,
                    value=f"=SUM({_col_letter(C_SUM_CONTR_NET)}{first}:"
                          f"{_col_letter(C_SUM_CONTR_NET)}{last})")
            _money_format(ws.cell(row=row, column=C_SUM_CONTR_NET))
            ws.cell(row=row, column=C_MARGIN_ABS,
                    value=f"={_col_letter(C_SUM_CLIENT_NET)}{row}"
                          f"-{_col_letter(C_SUM_CONTR_NET)}{row}")
            _money_format(ws.cell(row=row, column=C_MARGIN_ABS))
            ws.cell(row=row, column=C_MARGIN_PCT,
                    value=f"=IF({_col_letter(C_SUM_CLIENT_NET)}{row}=0,0,"
                          f"{_col_letter(C_MARGIN_ABS)}{row}"
                          f"/{_col_letter(C_SUM_CLIENT_NET)}{row})")
            _pct_format(ws.cell(row=row, column=C_MARGIN_PCT))

            section_total_rows.append((section, row, section_item_rows))
            row += 1
        row += 1  # пустая строка между разделами

    # Финальный итог проекта
    if section_total_rows:
        row += 1
        total_rows = [tr for _, tr, _ in section_total_rows]
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = PROJECT_TOTAL_FILL
            ws.cell(row=row, column=c).font = PROJECT_FONT
            ws.cell(row=row, column=c).border = BORDER_THIN
        ws.cell(row=row, column=1, value="ИТОГО ПО ПРОЕКТУ")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)

        for col_target in (C_SUM_CLIENT_GROSS, C_SUM_CLIENT_NET, C_SUM_CONTR_NET):
            parts = "+".join(f"{_col_letter(col_target)}{tr}" for tr in total_rows)
            ws.cell(row=row, column=col_target, value=f"={parts}")
            _money_format(ws.cell(row=row, column=col_target))

        ws.cell(row=row, column=C_MARGIN_ABS,
                value=f"={_col_letter(C_SUM_CLIENT_NET)}{row}"
                      f"-{_col_letter(C_SUM_CONTR_NET)}{row}")
        _money_format(ws.cell(row=row, column=C_MARGIN_ABS))
        ws.cell(row=row, column=C_MARGIN_PCT,
                value=f"=IF({_col_letter(C_SUM_CLIENT_NET)}{row}=0,0,"
                      f"{_col_letter(C_MARGIN_ABS)}{row}"
                      f"/{_col_letter(C_SUM_CLIENT_NET)}{row})")
        _pct_format(ws.cell(row=row, column=C_MARGIN_PCT))


# ---------------------------------------------------------- Аналитика
def _build_analytics_sheet(ws: Worksheet, client: Estimate, contractor: Estimate,
                            matches: list[Match]):
    ws.title = "Аналитика"
    for col, w in enumerate([42, 22, 22, 18, 14, 26], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    match_by_client = {m.client_idx: m for m in matches}

    # ----- Сбор данных
    work_by_section: dict[str, dict] = {}
    material_by_category: dict[str, dict] = {}

    for c_idx, it in enumerate(client.items):
        qty = it.quantity or 0
        match = match_by_client.get(c_idx)
        contractor_item = (
            contractor.items[match.contractor_idx]
            if match and match.contractor_idx is not None else None
        )

        # МАТЕРИАЛ
        if it.price_material is not None or it.sum_material is not None:
            sum_mat_gross = (it.sum_material if it.sum_material is not None
                             else (it.price_material or 0) * qty)
            sum_mat_net = sum_mat_gross / (1 + VAT_RATE)
            category = classify_material(it.name)
            d = material_by_category.setdefault(
                category, {"client_net": 0.0, "purchase_net": 0.0,
                           "items": 0, "with_price": 0}
            )
            d["client_net"] += sum_mat_net
            d["items"] += 1
            # Цена закупки материалов пока не приходит — оставляем 0 + помечаем

        # РАБОТА
        if it.price_work is not None or it.sum_work is not None:
            sum_work_gross = (it.sum_work if it.sum_work is not None
                              else (it.price_work or 0) * qty)
            sum_work_net = sum_work_gross / (1 + VAT_RATE)
            s = work_by_section.setdefault(
                it.section, {"client_net": 0.0, "contractor_net": 0.0,
                             "items": 0, "matched": 0}
            )
            s["client_net"] += sum_work_net
            s["items"] += 1
            if contractor_item:
                s["contractor_net"] += (contractor_item.price_work or 0) * qty
                s["matched"] += 1

    # ----- Шапка
    row = 1
    ws.cell(row=row, column=1, value="📊 Аналитика проекта").font = Font(bold=True, size=14)
    row += 2

    # ----- Метрики проекта
    total_client_net = sum(s["client_net"] for s in work_by_section.values()) + \
                       sum(m["client_net"] for m in material_by_category.values())
    total_contractor_net = sum(s["contractor_net"] for s in work_by_section.values())

    # «закрытая» маржа — только по работам с матчами
    closed_client_net = sum(
        s["client_net"] for s in work_by_section.values() if s["matched"] > 0
    )
    closed_margin = closed_client_net - total_contractor_net
    closed_margin_pct = closed_margin / closed_client_net if closed_client_net else None

    metrics = [
        ("Сумма проекта у клиента (без НДС, при НДС 22%)", total_client_net, "#,##0.00 ₽", None),
        ("Сумма закупки работ у подрядчика (без НДС)", total_contractor_net, "#,##0.00 ₽", None),
        ("Маржа по закрытым работам (руб.)", closed_margin, "#,##0.00 ₽", None),
        ("Маржа по закрытым работам, %", closed_margin_pct, "0.0%", closed_margin_pct),
        ("Позиций у клиента", len(client.items), "0", None),
        ("Из них с матчем подрядчика", sum(1 for m in matches if m.contractor_idx is not None), "0", None),
    ]
    for label, value, fmt, hue in metrics:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=row, column=2, value=value)
        c.number_format = fmt
        if hue is not None:
            c.fill = _margin_fill(hue)
        for col in (1, 2):
            ws.cell(row=row, column=col).border = BORDER_THIN
        row += 1
    row += 2

    # ----- Таблица: маржа по РАБОТАМ
    ws.cell(row=row, column=1, value="🔧 Работы — маржа по разделам").font = Font(bold=True, size=12)
    row += 1
    headers = ["Раздел", "Клиент (без НДС)", "Подрядчик (без НДС)",
               "Маржа, руб.", "Маржа, %", "Статус"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER_THIN
    row += 1

    work_first = row
    for s_name, data in work_by_section.items():
        ws.cell(row=row, column=1, value=s_name)
        ws.cell(row=row, column=2, value=data["client_net"]).number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=3, value=data["contractor_net"]).number_format = "#,##0.00 ₽"

        margin_formula = f"=B{row}-C{row}"
        ws.cell(row=row, column=4, value=margin_formula).number_format = "#,##0.00 ₽"
        pct_cell = ws.cell(row=row, column=5,
                            value=f"=IF(B{row}=0,0,D{row}/B{row})")
        pct_cell.number_format = "0.0%"

        if data["matched"] == 0:
            status = "⚠ Нет нижних цен"
            pct_cell.fill = GREY
        elif data["matched"] < data["items"]:
            status = f"Частично: {data['matched']} из {data['items']}"
            margin_val = (data["client_net"] - data["contractor_net"]) / data["client_net"] \
                if data["client_net"] else None
            pct_cell.fill = _margin_fill(margin_val)
        else:
            status = f"Закрыто ✓ ({data['items']} поз.)"
            margin_val = (data["client_net"] - data["contractor_net"]) / data["client_net"] \
                if data["client_net"] else None
            pct_cell.fill = _margin_fill(margin_val)
        ws.cell(row=row, column=6, value=status)
        for c in range(1, 7):
            ws.cell(row=row, column=c).border = BORDER_THIN
        row += 1

    # итог по работам
    if work_first < row:
        last = row - 1
        for c in range(1, 7):
            ws.cell(row=row, column=c).fill = TOTAL_FILL
            ws.cell(row=row, column=c).font = Font(bold=True)
            ws.cell(row=row, column=c).border = BORDER_THIN
        ws.cell(row=row, column=1, value="Итого по работам")
        ws.cell(row=row, column=2, value=f"=SUM(B{work_first}:B{last})").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=3, value=f"=SUM(C{work_first}:C{last})").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=4, value=f"=B{row}-C{row}").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=5,
                value=f"=IF(B{row}=0,0,D{row}/B{row})").number_format = "0.0%"
        row += 1
    row += 2

    # ----- Таблица: маржа по МАТЕРИАЛАМ (по категориям)
    ws.cell(row=row, column=1,
            value="📦 Материалы — по категориям").font = Font(bold=True, size=12)
    row += 1
    headers = ["Категория", "Клиент (без НДС)", "Закупка (без НДС)",
               "Маржа, руб.", "Маржа, %", "Статус"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER_THIN
    row += 1

    mat_first = row
    category_order = ["Лотки", "Кабель и проводка", "Зарядные станции",
                      "Шкафы и щиты", "Другое"]
    for category in category_order:
        if category not in material_by_category:
            continue
        data = material_by_category[category]
        ws.cell(row=row, column=1, value=category)
        ws.cell(row=row, column=2, value=data["client_net"]).number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=3, value=data["purchase_net"]).number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=4, value=f"=B{row}-C{row}").number_format = "#,##0.00 ₽"
        pct_cell = ws.cell(row=row, column=5,
                            value=f"=IF(B{row}=0,0,D{row}/B{row})")
        pct_cell.number_format = "0.0%"
        pct_cell.fill = GREY
        ws.cell(row=row, column=6,
                value=f"⚠ Нет цен закупки ({data['items']} поз.)").font = Font(italic=True)
        for c in range(1, 7):
            ws.cell(row=row, column=c).border = BORDER_THIN
        row += 1

    if mat_first < row:
        last = row - 1
        for c in range(1, 7):
            ws.cell(row=row, column=c).fill = TOTAL_FILL
            ws.cell(row=row, column=c).font = Font(bold=True)
            ws.cell(row=row, column=c).border = BORDER_THIN
        ws.cell(row=row, column=1, value="Итого по материалам")
        ws.cell(row=row, column=2,
                value=f"=SUM(B{mat_first}:B{last})").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=3,
                value=f"=SUM(C{mat_first}:C{last})").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=4, value=f"=B{row}-C{row}").number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=5,
                value=f"=IF(B{row}=0,0,D{row}/B{row})").number_format = "0.0%"


# ---------------------------------------------------------- Исходники с итогами
def _build_raw_sheet(ws: Worksheet, estimate: Estimate, title: str):
    ws.title = title[:31]
    headers = ["Раздел", "№", "Тип", "Наименование", "Ед.", "Кол-во",
               "Цена мат.", "Цена раб.", "Сумма мат.", "Сумма раб.", "Итого"]
    widths = [28, 5, 12, 55, 8, 10, 14, 14, 18, 18, 18]
    ncols = len(headers)
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # Группируем по разделам с итогами
    sections_in_order: list[str] = []
    seen = set()
    for it in estimate.items:
        if it.section not in seen:
            sections_in_order.append(it.section)
            seen.add(it.section)

    row = 2
    section_total_rows: list[int] = []

    for section in sections_in_order:
        # Заголовок раздела
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = SECTION_FILL
            ws.cell(row=row, column=c).border = BORDER_THIN
        ws.cell(row=row, column=1, value=section).font = Font(bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        row += 1

        item_rows = []
        for it in estimate.items:
            if it.section != section:
                continue
            ws.cell(row=row, column=1, value=it.section)
            ws.cell(row=row, column=2, value=it.number)
            ws.cell(row=row, column=3, value=it.kind)
            ws.cell(row=row, column=4, value=it.name).alignment = \
                Alignment(wrap_text=True, vertical="top")
            ws.cell(row=row, column=5, value=it.unit)
            ws.cell(row=row, column=6, value=it.quantity).number_format = "#,##0.00"

            qty_ref = f"F{row}"
            # Цены и суммы — формулы где возможно
            if it.price_material is not None:
                ws.cell(row=row, column=7, value=it.price_material).number_format = "#,##0.00"
            if it.price_work is not None:
                ws.cell(row=row, column=8, value=it.price_work).number_format = "#,##0.00"
            # Сумма мат. = qty * price_mat если оба есть, иначе ставим значение
            if it.price_material is not None and it.quantity is not None:
                ws.cell(row=row, column=9, value=f"={qty_ref}*G{row}").number_format = "#,##0.00"
            elif it.sum_material is not None:
                ws.cell(row=row, column=9, value=it.sum_material).number_format = "#,##0.00"
            if it.price_work is not None and it.quantity is not None:
                ws.cell(row=row, column=10, value=f"={qty_ref}*H{row}").number_format = "#,##0.00"
            elif it.sum_work is not None:
                ws.cell(row=row, column=10, value=it.sum_work).number_format = "#,##0.00"
            # Итого = сумма мат + сумма раб
            ws.cell(row=row, column=11,
                    value=f"=IFERROR(I{row},0)+IFERROR(J{row},0)").number_format = "#,##0.00"

            for c in range(1, ncols + 1):
                ws.cell(row=row, column=c).border = BORDER_THIN
            item_rows.append(row)
            row += 1

        # Итог по разделу
        if item_rows:
            first = item_rows[0]
            last = item_rows[-1]
            for c in range(1, ncols + 1):
                ws.cell(row=row, column=c).fill = TOTAL_FILL
                ws.cell(row=row, column=c).font = Font(bold=True)
                ws.cell(row=row, column=c).border = BORDER_THIN
            ws.cell(row=row, column=1, value=f"Итого: {section}")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row=row, column=9,
                    value=f"=SUM(I{first}:I{last})").number_format = "#,##0.00"
            ws.cell(row=row, column=10,
                    value=f"=SUM(J{first}:J{last})").number_format = "#,##0.00"
            ws.cell(row=row, column=11,
                    value=f"=SUM(K{first}:K{last})").number_format = "#,##0.00"
            section_total_rows.append(row)
            row += 1
        row += 1  # пустая строка

    # Финальный итог
    if section_total_rows:
        row += 1
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = PROJECT_TOTAL_FILL
            ws.cell(row=row, column=c).font = PROJECT_FONT
            ws.cell(row=row, column=c).border = BORDER_THIN
        ws.cell(row=row, column=1, value=f"ИТОГО ПО СМЕТЕ: {estimate.title}")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        for col, letter in [(9, "I"), (10, "J"), (11, "K")]:
            parts = "+".join(f"{letter}{tr}" for tr in section_total_rows)
            ws.cell(row=row, column=col, value=f"={parts}").number_format = "#,##0.00"


# ---------------------------------------------------------- Сборка
def build_workbook(client: Estimate, contractor: Estimate,
                    matches: list[Match]) -> io.BytesIO:
    wb = Workbook()
    _build_summary_sheet(wb.active, client, contractor, matches)
    _build_analytics_sheet(wb.create_sheet(), client, contractor, matches)
    _build_raw_sheet(wb.create_sheet(), client, "Исходник клиента")
    _build_raw_sheet(wb.create_sheet(), contractor, "Исходник подрядчика")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def filename_for(client: Estimate) -> str:
    safe = "".join(c if c.isalnum() or c in " -_." else "_" for c in client.title)
    return f"Сводная смета — {safe} — {date.today().isoformat()}.xlsx"
