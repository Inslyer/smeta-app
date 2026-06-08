"""AI-сопоставление позиций сметы клиента с позициями подрядчика через Claude.

Стратегия:
    1. Один батч-запрос со всеми работами клиента и всеми работами подрядчика.
    2. Claude возвращает JSON с парами (client_idx → contractor_idx | null),
       уверенностью и обоснованием.
    3. Используем tool_use чтобы гарантировать структурированный ответ
       без галлюцинаций по формату.

Кэширование: ответы сохраняются в .cache/matches/<hash>.json — повторный
прогон тех же двух смет не тратит токены.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Iterable

from anthropic import Anthropic

from .models import Estimate, EstimateItem, MaterialMatch, Match


CACHE_DIR = Path(__file__).resolve().parent.parent / ".cache" / "matches"
DEFAULT_MODEL = "claude-sonnet-4-6"


SYSTEM_PROMPT = """\
Ты — эксперт в чтении строительно-монтажных смет (электромонтаж, кабельные \
трассы, электрощитовое оборудование). Тебе дают список работ из верхней \
сметы (клиент) и список работ из нижней сметы подрядчика. Твоя задача — \
сопоставить КАЖДУЮ позицию клиента с РОВНО ОДНОЙ позицией подрядчика \
(или явно сказать что соответствия нет).

КРИТЕРИИ СОПОСТАВЛЕНИЯ:
1. Сечение/типоразмер кабеля должны совпадать. Кабель 5х70 ≠ кабель 5х95.
2. Тип работы должен совпадать: "прокладка кабеля" ≠ "монтаж муфты".
3. Бренд/марка кабеля (ППГнг-LS vs ППНнг-HF) — НЕ критичны для сопоставления \
работы по монтажу, но всё равно отмечай отличие в reason.
4. Лотки: размер (400х80, 200х80, 80х80) обязан совпадать.
5. Шкафы: сопоставляются по названию (ЩСУ-ЭЗС → ЩСУ ЭЗС).
6. Если работа клиента не покрыта подрядчиком — верни contractor_idx: null.
7. Одной позиции подрядчика разрешается соответствовать НЕСКОЛЬКИМ позициям \
клиента (например, монтаж лотка 200х80 может быть в нескольких местах).

ШКАЛА УВЕРЕННОСТИ:
- 0.95-1.00 — однозначное совпадение по типу и параметрам.
- 0.70-0.94 — совпадение с оговорками (бренд кабеля отличается, разные ед.изм., …).
- 0.40-0.69 — спорное совпадение — пометь в reason что нужно подтвердить.
- 0.00-0.39 — лучше вернуть null.

Возвращай ответ ТОЛЬКО через инструмент `submit_matches`. Никакого свободного \
текста до или после.
"""


MATCH_TOOL = {
    "name": "submit_matches",
    "description": (
        "Передать список сопоставлений между позициями сметы клиента "
        "и позициями сметы подрядчика."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "matches": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "client_idx": {
                            "type": "integer",
                            "description": "Индекс позиции клиента в переданном списке.",
                        },
                        "contractor_idx": {
                            "type": ["integer", "null"],
                            "description": (
                                "Индекс позиции подрядчика, либо null, "
                                "если соответствия нет."
                            ),
                        },
                        "confidence": {
                            "type": "number",
                            "minimum": 0,
                            "maximum": 1,
                        },
                        "reason": {
                            "type": "string",
                            "description": "Краткое обоснование на русском, 1-2 предложения.",
                        },
                    },
                    "required": ["client_idx", "contractor_idx", "confidence", "reason"],
                },
            }
        },
        "required": ["matches"],
    },
}


def _client_work_items(estimate: Estimate) -> list[tuple[int, EstimateItem]]:
    """Возвращает (исходный_индекс, item) только для тех позиций клиента,
    у которых есть работа (work / mixed / composite — у composite сам шкаф
    может иметь монтаж)."""
    return [
        (i, it)
        for i, it in enumerate(estimate.items)
        if it.kind in ("work", "mixed", "composite") and it.price_work
    ]


def _contractor_work_items(estimate: Estimate) -> list[tuple[int, EstimateItem]]:
    """У подрядчика всё — работы."""
    return [(i, it) for i, it in enumerate(estimate.items)]


def _format_items(pairs: Iterable[tuple[int, EstimateItem]]) -> str:
    lines = []
    for orig_idx, it in pairs:
        lines.append(
            f"[{orig_idx}] раздел: {it.section} | {it.name} | "
            f"{it.quantity} {it.unit}"
        )
    return "\n".join(lines)


def _cache_key(client: Estimate, contractor: Estimate, model: str) -> str:
    hasher = hashlib.sha256()
    hasher.update(model.encode())
    for est in (client, contractor):
        for it in est.items:
            hasher.update(
                f"{it.section}|{it.name}|{it.unit}|{it.quantity}|{it.kind}".encode()
            )
    return hasher.hexdigest()[:16]


def match_estimates(
    client: Estimate,
    contractor: Estimate,
    *,
    model: str | None = None,
    use_cache: bool = True,
) -> tuple[list[Match], dict]:
    """Главная функция сопоставления.

    Returns:
        (matches, meta) — meta содержит usage/from_cache.
    """
    model = model or os.getenv("CLAUDE_MODEL") or DEFAULT_MODEL

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"{_cache_key(client, contractor, model)}.json"

    if use_cache and cache_path.exists():
        with cache_path.open(encoding="utf-8") as f:
            payload = json.load(f)
        matches = [Match(**m) for m in payload["matches"]]
        return matches, {"from_cache": True, **payload.get("meta", {})}

    client_pairs = _client_work_items(client)
    contractor_pairs = _contractor_work_items(contractor)

    if not client_pairs or not contractor_pairs:
        return [], {"from_cache": False, "skipped": "empty"}

    user_prompt = (
        "СМЕТА КЛИЕНТА — позиции с работой:\n"
        f"{_format_items(client_pairs)}\n\n"
        "СМЕТА ПОДРЯДЧИКА — все позиции:\n"
        f"{_format_items(contractor_pairs)}\n\n"
        "Сопоставь каждую позицию клиента с одной позицией подрядчика "
        "(или верни contractor_idx=null). Используй РОВНО те индексы, что "
        "указаны в квадратных скобках."
    )

    client_api = Anthropic()
    response = client_api.messages.create(
        model=model,
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        tools=[MATCH_TOOL],
        tool_choice={"type": "tool", "name": "submit_matches"},
        messages=[{"role": "user", "content": user_prompt}],
    )

    tool_use = next(
        (b for b in response.content if b.type == "tool_use"),
        None,
    )
    if tool_use is None:
        raise RuntimeError(
            "Модель не вызвала инструмент submit_matches. "
            f"Ответ: {response.content!r}"
        )

    raw_matches = tool_use.input.get("matches", [])
    matches = [
        Match(
            client_idx=m["client_idx"],
            contractor_idx=m["contractor_idx"],
            confidence=float(m["confidence"]),
            reason=m["reason"],
        )
        for m in raw_matches
    ]

    meta = {
        "from_cache": False,
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
        "model": model,
    }

    with cache_path.open("w", encoding="utf-8") as f:
        json.dump(
            {
                "matches": [m.__dict__ for m in matches],
                "meta": meta,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    return matches, meta


# ============================================================================
# Сопоставление МАТЕРИАЛОВ клиента с позициями счетов поставщиков
# ============================================================================

MATERIALS_SYSTEM_PROMPT = """\
Ты — эксперт по электротехническим материалам и слаботочке. Тебе дают \
список МАТЕРИАЛОВ из сметы клиента и список позиций из счетов поставщиков \
(один или несколько счетов от закупщика). Твоя задача — сопоставить \
каждую позицию материала клиента с РОВНО ОДНОЙ позицией из счетов \
(или вернуть invoice_item_id: null, если ничего не подходит).

КРИТЕРИИ СОПОСТАВЛЕНИЯ (в порядке убывания приоритета):

1. **Артикул** — если в позиции клиента указан артикул и в позиции счёта \
есть точно такой же `article_supplier` или `article_manufacturer` — это \
жёсткое совпадение, confidence ≥ 0.95, match_kind="article".

2. **Точный типоразмер**:
   - Кабель: сечение и количество жил ОБЯЗАНЫ совпадать (5х70 ≠ 5х95).
   - Лоток: ширина и высота ОБЯЗАНЫ совпадать (400х80 ≠ 200х80).
   - Профиль / шпилька: длина и сечение должны совпадать.
   - Шкаф / щит: исполнение и габариты.
   - Метизы (винты, гайки, шайбы): резьба и тип (М6, М8 и т. д.).

3. **Единица измерения** должна быть совместима: м ↔ м, шт ↔ шт. \
Если у клиента «компл», а в счёте «шт» — это нестрашно, но отметь \
в reason.

4. **Бренд / производитель** — НЕ критичен. Лоток одного завода \
заменяется лотком другого, если типоразмер совпадает. Отметь в reason.

5. **Несколько подходящих позиций счёта** — выбери ту, у которой \
типоразмер точнее совпадает. Если есть несколько счетов с одинаковым \
товаром — бери позицию с более свежей датой счёта.

6. Если позиция клиента — НЕ материал (это работа), верни \
invoice_item_id: null с confidence: 0.

ШКАЛА УВЕРЕННОСТИ:
- 0.95-1.00 — совпадение по артикулу или 1-в-1 по типоразмеру.
- 0.70-0.94 — типоразмер совпадает, бренд/производитель отличается.
- 0.40-0.69 — спорное совпадение — отметь в reason, что нужно подтвердить.
- 0.00-0.39 — лучше вернуть null.

match_kind:
- "article" — совпали артикулы.
- "semantic" — совпали по смыслу (типоразмер, наименование).
- "none" — соответствия нет (invoice_item_id = null).

Возвращай ответ ТОЛЬКО через инструмент `submit_material_matches`. Никакого \
свободного текста до или после.
"""


MATERIALS_MATCH_TOOL = {
    "name": "submit_material_matches",
    "description": (
        "Передать список сопоставлений позиций материалов клиента с "
        "позициями счетов поставщиков."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "matches": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "client_idx": {
                            "type": "integer",
                            "description": "Индекс материала клиента в переданном списке.",
                        },
                        "invoice_item_id": {
                            "type": ["integer", "null"],
                            "description": (
                                "id позиции счёта из колонки 'id' в переданной "
                                "таблице, либо null если соответствия нет."
                            ),
                        },
                        "confidence": {
                            "type": "number",
                            "minimum": 0,
                            "maximum": 1,
                        },
                        "reason": {
                            "type": "string",
                            "description": "Краткое обоснование на русском, 1-2 предложения.",
                        },
                        "match_kind": {
                            "type": "string",
                            "enum": ["article", "semantic", "none"],
                        },
                    },
                    "required": [
                        "client_idx", "invoice_item_id",
                        "confidence", "reason", "match_kind",
                    ],
                },
            },
        },
        "required": ["matches"],
    },
}


def _client_material_items(
    estimate: Estimate,
) -> list[tuple[int, EstimateItem]]:
    """Возвращает (исходный_индекс, item) только для материалов клиента."""
    return [
        (i, it)
        for i, it in enumerate(estimate.items)
        if it.kind in ("material", "mixed", "composite") and (it.price_material or it.sum_material)
    ]


def _format_client_materials(pairs: Iterable[tuple[int, EstimateItem]]) -> str:
    lines = []
    for orig_idx, it in pairs:
        lines.append(
            f"[{orig_idx}] раздел: {it.section} | {it.name} | "
            f"{it.quantity} {it.unit}"
        )
    return "\n".join(lines)


def _format_supplier_items(rows: list[dict]) -> str:
    """rows — это dict-и из db.all_invoice_items_for_project()."""
    lines = []
    for r in rows:
        art_s = r.get("article_supplier") or ""
        art_m = r.get("article_manufacturer") or ""
        articles = "/".join(x for x in (art_s, art_m) if x) or "—"
        unit = r.get("unit") or ""
        qty = r.get("quantity") or ""
        price = r.get("unit_price") or ""
        sup = r.get("supplier_name") or ""
        inv = r.get("invoice_number") or ""
        dt = r.get("invoice_date") or ""
        lines.append(
            f"id={r['id']} | арт: {articles} | {r.get('name','')} | "
            f"{qty} {unit} @ {price} | {sup} счёт №{inv} от {dt}"
        )
    return "\n".join(lines)


def _materials_cache_key(client: Estimate, supplier_rows: list[dict],
                          model: str) -> str:
    hasher = hashlib.sha256()
    hasher.update(model.encode())
    hasher.update(b"materials::")
    for it in client.items:
        hasher.update(
            f"{it.section}|{it.name}|{it.unit}|{it.quantity}|{it.kind}".encode()
        )
    hasher.update(b"::supplier::")
    for r in supplier_rows:
        hasher.update(
            f"{r.get('id')}|{r.get('article_supplier')}|{r.get('article_manufacturer')}|"
            f"{r.get('name')}|{r.get('unit_price')}".encode()
        )
    return hasher.hexdigest()[:16]


def match_materials(
    client: Estimate,
    supplier_rows: list[dict],
    *,
    model: str | None = None,
    use_cache: bool = True,
) -> tuple[list[MaterialMatch], dict]:
    """Сопоставляет материалы из сметы клиента с позициями счетов поставщиков.

    Args:
        client: смета клиента.
        supplier_rows: позиции всех счетов проекта (как dict-и из
            db.all_invoice_items_for_project), должны содержать поля
            id, name, article_supplier, article_manufacturer, unit,
            quantity, unit_price, supplier_name, invoice_number, invoice_date.

    Returns:
        (matches, meta) — meta содержит usage/from_cache.
    """
    model = model or os.getenv("CLAUDE_MODEL") or DEFAULT_MODEL

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / (
        f"materials_{_materials_cache_key(client, supplier_rows, model)}.json"
    )

    if use_cache and cache_path.exists():
        with cache_path.open(encoding="utf-8") as f:
            payload = json.load(f)
        matches = [MaterialMatch(**m) for m in payload["matches"]]
        return matches, {"from_cache": True, **payload.get("meta", {})}

    client_pairs = _client_material_items(client)
    if not client_pairs:
        return [], {"from_cache": False, "skipped": "no_client_materials"}
    if not supplier_rows:
        return [], {"from_cache": False, "skipped": "no_supplier_items"}

    user_prompt = (
        "СМЕТА КЛИЕНТА — материалы:\n"
        f"{_format_client_materials(client_pairs)}\n\n"
        "ПОЗИЦИИ ИЗ СЧЕТОВ ПОСТАВЩИКОВ (этот проект):\n"
        f"{_format_supplier_items(supplier_rows)}\n\n"
        "Сопоставь каждую позицию материала клиента с одной позицией из "
        "счетов (или верни invoice_item_id=null). Используй РОВНО те id, "
        "что указаны в поле id= у позиций счёта. Используй РОВНО те "
        "индексы в квадратных скобках для client_idx."
    )

    client_api = Anthropic()
    response = client_api.messages.create(
        model=model,
        max_tokens=8000,
        system=MATERIALS_SYSTEM_PROMPT,
        tools=[MATERIALS_MATCH_TOOL],
        tool_choice={"type": "tool", "name": "submit_material_matches"},
        messages=[{"role": "user", "content": user_prompt}],
    )

    tool_use = next(
        (b for b in response.content if b.type == "tool_use"), None,
    )
    if tool_use is None:
        raise RuntimeError(
            "Модель не вызвала submit_material_matches. "
            f"Ответ: {response.content!r}"
        )

    raw = tool_use.input.get("matches", [])
    matches = [
        MaterialMatch(
            client_idx=m["client_idx"],
            invoice_item_id=m["invoice_item_id"],
            confidence=float(m["confidence"]),
            reason=m["reason"],
            match_kind=m.get("match_kind", "semantic"),
        )
        for m in raw
    ]

    meta = {
        "from_cache": False,
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
        "model": model,
    }
    with cache_path.open("w", encoding="utf-8") as f:
        json.dump(
            {"matches": [m.__dict__ for m in matches], "meta": meta},
            f, ensure_ascii=False, indent=2,
        )
    return matches, meta
