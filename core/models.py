"""Модели данных для смет."""
from dataclasses import dataclass, field
from typing import Optional, Literal


ItemKind = Literal["material", "work", "mixed", "composite"]
Source = Literal["client", "contractor", "supplier"]


@dataclass
class EstimateItem:
    """Одна позиция в смете."""
    section: str
    number: str
    name: str
    unit: str
    quantity: float
    kind: ItemKind = "material"

    price_material: Optional[float] = None
    price_work: Optional[float] = None
    sum_material: Optional[float] = None
    sum_work: Optional[float] = None
    sum_total: Optional[float] = None

    vat_included: bool = True
    raw_row_index: Optional[int] = None
    components: list["EstimateItem"] = field(default_factory=list)


@dataclass
class Estimate:
    """Смета целиком."""
    source: Source
    title: str
    file_name: str
    items: list[EstimateItem] = field(default_factory=list)
    vat_included: bool = True
    total: Optional[float] = None


@dataclass
class Match:
    """AI-сопоставление позиции клиента с позицией подрядчика (или поставщика)."""
    client_idx: int                       # индекс в Estimate.items клиента
    contractor_idx: Optional[int]         # индекс позиции подрядчика; None = нет соответствия
    confidence: float                      # 0..1
    reason: str                            # обоснование от модели
    confirmed: bool = False                # подтверждено ли пользователем вручную
