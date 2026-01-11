
"""
templates.py

Paper sizes and basic layout rules for drawing sheets.
All units: points (1/72 inch).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Tuple

from reportlab.lib.pagesizes import A4, A3, landscape, portrait


@dataclass(frozen=True)
class SheetTemplate:
    name: str
    pagesize: Tuple[float, float]  # (width_pt, height_pt)


TEMPLATES: Dict[str, SheetTemplate] = {
    "A4_Landscape": SheetTemplate("A4_Landscape", landscape(A4)),
    "A4_Portrait": SheetTemplate("A4_Portrait", portrait(A4)),
    "A3_Landscape": SheetTemplate("A3_Landscape", landscape(A3)),
    "A3_Portrait": SheetTemplate("A3_Portrait", portrait(A3)),
}


def get_template_names() -> list[str]:
    return list(TEMPLATES.keys())


def get_template(name: str) -> SheetTemplate:
    return TEMPLATES.get(name, TEMPLATES["A3_Landscape"])
