"""Генератор итогового Excel со сводной сметой и аналитикой.

Структура книги:
    1. «Сводная смета» — позиции клиента, цены подрядчика, маржа по строке,
        итоги по разделам, итог проекта.
    2. «Аналитика» — общие метрики, маржа по разделам.
    3. «Исходник клиента» / «Исходник подрядчика» — оригинальные сметы.

Маржа считается в ценах БЕЗ НДС (клиент с НДС → делится на 1+VAT_RATE).
"""
from __future__ import annotations

import io
from datetime import date
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import Estimate, EstimateItem, Match


VAT_RATE = 0.20  # НДС клиента; подрядчик считается без НДС

# Палитра светофора по марже %.
GREEN = PatternFill("solid", fgColor="C6EFCE")     # маржа > 25%
YELLOW = PatternFill("solid", fgColor="FFEB9C")    # 10-25%
RED = PatternFill("solid", fgColor="FFC7CE")       # < 10% или отрицательная
GREY = PatternFill("solid", fgColor="EEECEC")      # нет данных у подрядчика
SECTION_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="B4C7E7")
PROJECT_TOTAL_FILL = PatternFill("solid", fgColor="2E75B6")

BORDER_THIN = Border(
    left=Side(style="thin", color="C0C0C0"),
    right=Side(style="thin", color="C0C0C0"),
    top=Side(style="thin", color="C0C0C0"),
    bottom=Side(style="thin", color="C0C0C0"),
)

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")


def _margin_fill(margin_pct: Optional[float]) -> PatternFill:
    if margin_pct is None:
        return GREY
    if margin_pct >= 0.25:
        return GREEN
    if margin_pct >= 0.10:
        return YELLOW
    return RED


def _net(value: Optional[float], vat_included: bool) -> Optional[float]:
    if value is None:
        return None
    if vat_included:
        return value / (1 + VAT_RATE)
    return value


def _safe_div(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None or b == 0:
        return None
    return a / b


# --------------------------------------------------------------------- Sheet 1
COLUMNS = [
    ("Раздел", 28),
    ("№", 5),
    ("Наименование", 55),
    ("Ед.", 8),
    ("Кол-во", 10),
    ("Цена клиента (с НДС)", 16),
    ("Сумма клиента (с НДС)", 18),
    ("Сумма клиента (без НДС)", 18),
    ("Исполнитель", 22),
    ("Цена подрядчика", 16),
    ("Сумма подрядчика (без НДС)", 20),
    ("Маржа, руб.", 14),
    ("Маржа, %", 11),
    ("Уверенность AI", 14),
    ("Комментарий AI", 50),
]


def _write_headers(ws):
    for col_idx, (title, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "C2"


def _section_row(ws, row_idx: int, section: str):
    cell = ws.cell(row=row_idx, column=1, value=section)
    cell.font = Font(bold=True)
    cell.fill = SECTION_FILL
    for c in range(1, len(COLUMNS) + 1):
        ws.cell(row=row_idx, column=c).fill = SECTION_FILL
        ws.cell(row=row_idx, column=c).border = BORDER_THIN
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(COLUMNS))


def _section_total_row(
    ws,
    row_idx: int,
    section: str,
    sum_client: float,
    sum_client_net: float,
    sum_contractor_net: float,
):
    margin = sum_client_net - sum_contractor_net
    margin_pct = _safe_div(margin, sum_client_net)
    label = f"Итого по разделу: {section}"
    ws.cell(row=row_idx, column=1, value=label).font = Font(bold=True)
    ws.cell(row=row_idx, column=7, value=sum_client).font = Font(bold=True)
    ws.cell(row=row_idx, column=8, value=sum_client_net).font = Font(bold=True)
    ws.cell(row=row_idx, column=11, value=sum_contractor_net).font = Font(bold=True)
    ws.cell(row=row_idx, column=12, value=margin).font = Font(bold=True)
    pct_cell = ws.cell(row=row_idx, column=13, value=margin_pct)
    pct_cell.font = Font(bold=True)
    pct_cell.number_format = "0.0%"
    pct_cell.fill = _margin_fill(margin_pct)
    for c in range(1, len(COLUMNS) + 1):
        ws.cell(row=row_idx, column=c).border = BORDER_THIN
        if not ws.cell(row=row_idx, column=c).fill or \
           ws.cell(row=row_idx, column=c).fill.fgColor.rgb == "00000000":
            ws.cell(row=row_idx, column=c).fill = TOTAL_FILL
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)


def _project_total_row(
    ws,
    row_idx: int,
    sum_client: float,
    sum_client_net: float,
    sum_contractor_net: float,
):
    margin = sum_client_net - sum_contractor_net
    margin_pct = _safe_div(margin, sum_client_net)
    for c in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = PROJECT_TOTAL_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = BORDER_THIN
    ws.cell(row=row_idx, column=1, value="ИТОГО ПО ПРОЕКТУ")
    ws.cell(row=row_idx, column=7, value=sum_client).number_format = "#,##0.00 ₽"
    ws.cell(row=row_idx, column=8, value=sum_client_net).number_format = "#,##0.00 ₽"
    ws.cell(row=row_idx, column=11, value=sum_contractor_net).number_format = "#,##0.00 ₽"
    ws.cell(row=row_idx, column=12, value=margin).number_format = "#,##0.00 ₽"
    pct_cell = ws.cell(row=row_idx, column=13, value=margin_pct)
    pct_cell.number_format = "0.0%"
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)


def _write_item_row(
    ws,
    row_idx: int,
    item: EstimateItem,
    match: Optional[Match],
    contractor: Estimate,
) -> tuple[float, float, float]:
    """Возвращает (sum_client_gross, sum_client_net, sum_contractor_net) — для итогов."""
    contractor_item = (
        contractor.items[match.contractor_idx]
        if match and match.contractor_idx is not None else None
    )

    sum_client_gross = item.sum_total or 0
    sum_client_net = _net(sum_client_gross, vat_included=item.vat_included) or 0

    sum_contractor_net = 0.0
    contractor_unit_price = None
    if contractor_item and item.kind in ("work", "mixed", "composite"):
        unit = contractor_item.price_work or 0
        contractor_unit_price = unit
        sum_contractor_net = unit * (item.quantity or 0)

    # клиент = материал + работа в одной строке; для маржи берём всю сумму "без НДС"
    margin = sum_client_net - sum_contractor_net
    margin_pct = _safe_div(margin, sum_client_net)

    unit_price_client = None
    if item.price_material is not None and item.price_work is not None:
        unit_price_client = item.price_material + item.price_work
    elif item.price_work is not None:
        unit_price_client = item.price_work
    elif item.price_material is not None:
        unit_price_client = item.price_material

    ws.cell(row=row_idx, column=1, value=item.section)
    ws.cell(row=row_idx, column=2, value=item.number)
    ws.cell(row=row_idx, column=3, value=item.name).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=row_idx, column=4, value=item.unit)
    ws.cell(row=row_idx, column=5, value=item.quantity)
    ws.cell(row=row_idx, column=6, value=unit_price_client)
    ws.cell(row=row_idx, column=7, value=sum_client_gross)
    ws.cell(row=row_idx, column=8, value=sum_client_net)
    ws.cell(row=row_idx, column=9,
            value=("Подрядчик: " + contractor.title) if contractor_item else "— не закрыто —")
    ws.cell(row=row_idx, column=10, value=contractor_unit_price)
    ws.cell(row=row_idx, column=11, value=sum_contractor_net if contractor_item else None)
    ws.cell(row=row_idx, column=12, value=margin if contractor_item else None)
    pct_cell = ws.cell(row=row_idx, column=13, value=margin_pct if contractor_item else None)
    pct_cell.number_format = "0.0%"

    if match:
        ws.cell(row=row_idx, column=14, value=match.confidence).number_format = "0.0%"
        ws.cell(row=row_idx, column=15, value=match.reason).alignment = \
            Alignment(wrap_text=True, vertical="top")

    # форматы чисел
    for col, fmt in [(5, "#,##0.00"), (6, "#,##0.00 ₽"), (7, "#,##0.00 ₽"),
                     (8, "#,##0.00 ₽"), (10, "#,##0.00 ₽"), (11, "#,##0.00 ₽"),
                     (12, "#,##0.00 ₽")]:
        ws.cell(row=row_idx, column=col).number_format = fmt

    pct_cell.fill = _margin_fill(margin_pct if contractor_item else None)
    for c in range(1, len(COLUMNS) + 1):
        ws.cell(row=row_idx, column=c).border = BORDER_THIN

    return sum_client_gross, sum_client_net, sum_contractor_net


def _build_summary_sheet(ws, client: Estimate, contractor: Estimate,
                         matches: list[Match]):
    """Сводная смета."""
    ws.title = "Сводная смета"
    _write_headers(ws)

    # быстрый доступ к матчам по client_idx
    match_by_client = {m.client_idx: m for m in matches}

    row_idx = 2
    project_client_gross = 0.0
    project_client_net = 0.0
    project_contractor_net = 0.0

    sections_in_order: list[str] = []
    seen = set()
    for it in client.items:
        if it.section not in seen:
            sections_in_order.append(it.section)
            seen.add(it.section)

    for section in sections_in_order:
        _section_row(ws, row_idx, section)
        row_idx += 1

        section_client_gross = 0.0
        section_client_net = 0.0
        section_contractor_net = 0.0

        for idx, item in enumerate(client.items):
            if item.section != section:
                continue
            m = match_by_client.get(idx)
            gross, net, c_net = _write_item_row(ws, row_idx, item, m, contractor)
            row_idx += 1
            section_client_gross += gross
            section_client_net += net
            section_contractor_net += c_net

        _section_total_row(
            ws, row_idx, section,
            section_client_gross, section_client_net, section_contractor_net,
        )
        row_idx += 1
        project_client_gross += section_client_gross
        project_client_net += section_client_net
        project_contractor_net += section_contractor_net

    row_idx += 1
    _project_total_row(
        ws, row_idx,
        project_client_gross, project_client_net, project_contractor_net,
    )


# --------------------------------------------------------------------- Sheet 2
def _build_analytics_sheet(ws, client: Estimate, contractor: Estimate,
                            matches: list[Match]):
    ws.title = "Аналитика"
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14

    match_by_client = {m.client_idx: m for m in matches}

    # агрегаты по разделам
    sections: dict[str, dict] = {}
    project = {"client_gross": 0.0, "client_net": 0.0, "contractor_net": 0.0}

    for idx, it in enumerate(client.items):
        s = sections.setdefault(it.section, {"client_gross": 0.0, "client_net": 0.0,
                                             "contractor_net": 0.0, "items": 0,
                                             "closed": 0})
        client_gross = it.sum_total or 0
        client_net = _net(client_gross, vat_included=it.vat_included) or 0
        s["client_gross"] += client_gross
        s["client_net"] += client_net
        s["items"] += 1
        project["client_gross"] += client_gross
        project["client_net"] += client_net

        m = match_by_client.get(idx)
        if m and m.contractor_idx is not None and it.kind in ("work", "mixed", "composite"):
            cp = contractor.items[m.contractor_idx]
            c_sum = (cp.price_work or 0) * (it.quantity or 0)
            s["contractor_net"] += c_sum
            s["closed"] += 1
            project["contractor_net"] += c_sum

    row = 1
    ws.cell(row=row, column=1, value="Метрика").font = HEADER_FONT
    ws.cell(row=row, column=2, value="Значение").font = HEADER_FONT
    for c in (1, 2):
        ws.cell(row=row, column=c).fill = HEADER_FILL
        ws.cell(row=row, column=c).border = BORDER_THIN
    row += 1

    project_margin = project["client_net"] - project["contractor_net"]
    project_margin_pct = _safe_div(project_margin, project["client_net"])

    metrics = [
        ("Сумма клиента (с НДС)", project["client_gross"], "#,##0.00 ₽"),
        ("Сумма клиента (без НДС)", project["client_net"], "#,##0.00 ₽"),
        ("Закупка / подрядчик (без НДС)", project["contractor_net"], "#,##0.00 ₽"),
        ("Маржа (руб., без НДС)", project_margin, "#,##0.00 ₽"),
        ("Маржа (%)", project_margin_pct, "0.0%"),
        ("Позиций у клиента", len(client.items), "0"),
        ("Из них закрыто матчем", sum(1 for m in matches if m.contractor_idx is not None), "0"),
        ("Уверенные матчи (≥90%)",
         sum(1 for m in matches if m.confidence >= 0.9 and m.contractor_idx is not None), "0"),
    ]
    for label, value, fmt in metrics:
        ws.cell(row=row, column=1, value=label)
        c = ws.cell(row=row, column=2, value=value)
        c.number_format = fmt
        if "Маржа (%)" in label:
            c.fill = _margin_fill(value)
        for col in (1, 2):
            ws.cell(row=row, column=col).border = BORDER_THIN
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="Раздел").font = HEADER_FONT
    ws.cell(row=row, column=2, value="Клиент (без НДС)").font = HEADER_FONT
    ws.cell(row=row, column=3, value="Закупка (без НДС)").font = HEADER_FONT
    ws.cell(row=row, column=4, value="Маржа %").font = HEADER_FONT
    for c in (1, 2, 3, 4):
        ws.cell(row=row, column=c).fill = HEADER_FILL
        ws.cell(row=row, column=c).border = BORDER_THIN
    row += 1

    for s_name, data in sections.items():
        margin_pct = _safe_div(
            data["client_net"] - data["contractor_net"], data["client_net"]
        )
        ws.cell(row=row, column=1, value=s_name)
        ws.cell(row=row, column=2, value=data["client_net"]).number_format = "#,##0.00 ₽"
        ws.cell(row=row, column=3, value=data["contractor_net"]).number_format = "#,##0.00 ₽"
        m_cell = ws.cell(row=row, column=4, value=margin_pct)
        m_cell.number_format = "0.0%"
        m_cell.fill = _margin_fill(margin_pct)
        for c in (1, 2, 3, 4):
            ws.cell(row=row, column=c).border = BORDER_THIN
        row += 1


# --------------------------------------------------------------------- Sheet 3+
def _build_raw_sheet(ws, estimate: Estimate, title: str):
    ws.title = title[:31]
    headers = ["Раздел", "№", "Тип", "Наименование", "Ед.", "Кол-во",
               "Цена мат.", "Цена раб.", "Сумма мат.", "Сумма раб.", "Итого"]
    widths = [28, 5, 12, 55, 8, 10, 14, 14, 16, 16, 16]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    row = 2
    for it in estimate.items:
        ws.cell(row=row, column=1, value=it.section)
        ws.cell(row=row, column=2, value=it.number)
        ws.cell(row=row, column=3, value=it.kind)
        ws.cell(row=row, column=4, value=it.name).alignment = \
            Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=5, value=it.unit)
        ws.cell(row=row, column=6, value=it.quantity)
        for col, val, fmt in [
            (7, it.price_material, "#,##0.00"),
            (8, it.price_work, "#,##0.00"),
            (9, it.sum_material, "#,##0.00"),
            (10, it.sum_work, "#,##0.00"),
            (11, it.sum_total, "#,##0.00"),
        ]:
            if val is not None:
                ws.cell(row=row, column=col, value=val).number_format = fmt
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).border = BORDER_THIN
        row += 1


# --------------------------------------------------------------------- Build
def build_workbook(
    client: Estimate,
    contractor: Estimate,
    matches: list[Match],
) -> io.BytesIO:
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
