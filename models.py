

"""
models.py

Dataclasses for global title block fields + export settings + per-sheet overrides.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class TitleBlock:
    # Issuer info
    issuer_company: str = ""
    logo_path: str = ""  # optional path to PNG/JPG logo

    # Global drawing metadata (applies to all sheets)
    project: str = ""
    client: str = ""
    drawing_number: str = ""
    revision: str = ""
    date: str = ""
    drawn_by: str = ""
    checked_by: str = ""
    approved_by: str = ""
    # Per-sheet fields (drawing_title/comments) will be overridden per sheet


@dataclass
class ExportSettings:
    template_name: str = "A3_Landscape"
    fit_mode: str = "fit"  # "fit" or "fill"

    page_margin_pt: float = 18.0
    title_block_width_pt: float = 210.0
    header_height_pt: float = 0.0


@dataclass
class SheetPlanItem:
    """
    One output sheet.

    kind:
      - "image": source_path is an image file
      - "pdf": source_path is a PDF file, pdf_page_index selects which page
    """
    kind: str  # "image" | "pdf"
    source_path: str
    source_label: str  # shown in GUI (eg "file.pdf - Page 3")
    pdf_page_index: Optional[int] = None  # 0-based for PDFs, None for images

    # Per-sheet overrides
    drawing_title: str = ""
    comments: str = ""

