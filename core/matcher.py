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

from .models import Estimate, EstimateItem, Match


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
