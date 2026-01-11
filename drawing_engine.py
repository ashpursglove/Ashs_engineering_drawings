
"""
drawing_engine.py

Exports a multi-page engineering drawing pack:
- Everything uses the same sheet template and viewport rules.
- ISO title block strip on right.
- Page numbering Sheet X of Y.
- PDFs are rasterized per-page (each page becomes one sheet).
- Per-sheet Drawing Title + Comments supported.

Dependencies:
- Pillow
- reportlab
- pymupdf (fitz)
"""

from __future__ import annotations

import os
from dataclasses import asdict
from typing import List, Tuple, Union

import fitz  # PyMuPDF
from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

from models import ExportSettings, SheetPlanItem, TitleBlock
from templates import get_template


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
PDF_EXTS = {".pdf"}

ImageSource = Union[str, Image.Image]


def export_sheet_plan_to_pdf(
    sheet_plan: List[SheetPlanItem],
    output_pdf_path: str,
    global_tb: TitleBlock,
    settings: ExportSettings,
    pdf_render_dpi: int = 220,
) -> None:
    """
    Export a prepared sheet plan to a single PDF.

    sheet_plan determines:
    - what each sheet's source is
    - per-sheet drawing title + comments
    """
    if not sheet_plan:
        raise ValueError("No sheets to export (sheet plan is empty).")

    tpl = get_template(settings.template_name)
    page_w, page_h = tpl.pagesize

    c = canvas.Canvas(output_pdf_path, pagesize=(page_w, page_h))
    c.setTitle(os.path.basename(output_pdf_path))

    total = len(sheet_plan)

    for idx, item in enumerate(sheet_plan, start=1):
        img_src = _sheet_item_to_image(item, dpi=pdf_render_dpi)

        tb_for_sheet = _compose_title_block_for_sheet(
            global_tb=global_tb,
            item=item,
        )

        _render_one_sheet(
            c=c,
            page_w=page_w,
            page_h=page_h,
            image_source=img_src,
            title_block=tb_for_sheet,
            settings=settings,
            sheet_no=idx,
            sheet_total=total,
        )
        c.showPage()

        # If we created a PIL image for this sheet, let it be GC'd quickly
        if isinstance(img_src, Image.Image):
            img_src.close()

    c.save()


def _sheet_item_to_image(item: SheetPlanItem, dpi: int) -> ImageSource:
    """
    Convert a sheet plan item to an image source (path for images, PIL.Image for PDFs).
    """
    kind = (item.kind or "").lower()
    if kind == "image":
        return item.source_path

    if kind == "pdf":
        if item.pdf_page_index is None:
            raise ValueError(f"PDF sheet plan item missing pdf_page_index: {item.source_label}")
        return _render_pdf_page_to_image(item.source_path, item.pdf_page_index, dpi=dpi)

    raise ValueError(f"Unknown sheet plan kind: {item.kind}")


def _render_pdf_page_to_image(pdf_path: str, page_index: int, dpi: int = 220) -> Image.Image:
    """
    Render one PDF page into a PIL Image.
    """
    if dpi <= 0:
        dpi = 220

    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise ValueError(f"PDF page_index out of range: {page_index} for {pdf_path}")

        page = doc.load_page(page_index)
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return img
    finally:
        doc.close()


def _compose_title_block_for_sheet(global_tb: TitleBlock, item: SheetPlanItem) -> TitleBlock:
    """
    Merge global fields + per-sheet overrides into a TitleBlock-like object
    used by the title block renderer.
    """
    tb = TitleBlock(**asdict(global_tb))

    # Stash per-sheet fields into attributes we pass through the renderer
    # (renderer expects drawing_title/comments to exist; we pass them as dynamic attributes)
    # We'll attach them in a safe way by monkey-patching attributes on the instance.
    # (Keeps TitleBlock dataclass clean for "global" fields.)

    # Default title if user didn't set one for the sheet
    if item.drawing_title.strip():
        setattr(tb, "drawing_title", item.drawing_title.strip())
    else:
        # If empty: use source_label base
        setattr(tb, "drawing_title", _default_title_from_label(item.source_label))

    setattr(tb, "comments", item.comments.strip())

    return tb


def _default_title_from_label(label: str) -> str:
    base = os.path.basename(label)
    # Remove extension if it looks like a file
    name, _ext = os.path.splitext(base)
    return name or base


# ---------------------------------------------------------------------
# Sheet rendering
# ---------------------------------------------------------------------

def _render_one_sheet(
    c: canvas.Canvas,
    page_w: float,
    page_h: float,
    image_source: ImageSource,
    title_block: TitleBlock,
    settings: ExportSettings,
    sheet_no: int,
    sheet_total: int,
) -> None:
    margin = settings.page_margin_pt
    tb_w = settings.title_block_width_pt
    header_h = settings.header_height_pt

    viewport_x = margin
    viewport_y = margin
    viewport_w = page_w - (margin * 2) - tb_w
    viewport_h = page_h - (margin * 2) - header_h

    # Viewport frame
    c.saveState()
    c.setLineWidth(0.8)
    c.rect(viewport_x, viewport_y, viewport_w, viewport_h, stroke=1, fill=0)
    c.restoreState()

    _draw_image_fitted(
        c=c,
        image_source=image_source,
        box=(viewport_x, viewport_y, viewport_w, viewport_h),
        fit_mode=settings.fit_mode,
        inner_pad_pt=6.0,
    )

    _draw_iso_title_block(
        c=c,
        page_w=page_w,
        page_h=page_h,
        title_block=title_block,
        settings=settings,
        sheet_no=sheet_no,
        sheet_total=sheet_total,
    )


def _draw_image_fitted(
    c: canvas.Canvas,
    image_source: ImageSource,
    box: Tuple[float, float, float, float],
    fit_mode: str = "fit",
    inner_pad_pt: float = 6.0,
) -> None:
    x, y, w, h = box

    x2 = x + inner_pad_pt
    y2 = y + inner_pad_pt
    w2 = max(1.0, w - 2 * inner_pad_pt)
    h2 = max(1.0, h - 2 * inner_pad_pt)

    if isinstance(image_source, Image.Image):
        im_w, im_h = image_source.size
        img_reader = ImageReader(image_source)
    else:
        with Image.open(image_source) as im:
            im_w, im_h = im.size
        img_reader = ImageReader(image_source)

    if im_w <= 0 or im_h <= 0:
        return

    img_aspect = im_w / im_h
    box_aspect = w2 / h2

    if fit_mode not in {"fit", "fill"}:
        fit_mode = "fit"

    if fit_mode == "fit":
        if img_aspect > box_aspect:
            draw_w = w2
            draw_h = w2 / img_aspect
        else:
            draw_h = h2
            draw_w = h2 * img_aspect
    else:
        if img_aspect > box_aspect:
            draw_h = h2
            draw_w = h2 * img_aspect
        else:
            draw_w = w2
            draw_h = w2 / img_aspect

    draw_x = x2 + (w2 - draw_w) / 2.0
    draw_y = y2 + (h2 - draw_h) / 2.0

    c.drawImage(img_reader, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")


# ---------------------------------------------------------------------
# Text wrapping helpers (robust, breaks long tokens)
# ---------------------------------------------------------------------

def _wrap_text_to_lines(
    c: canvas.Canvas,
    text: str,
    max_width: float,
    font_name: str,
    font_size: int,
) -> list[str]:
    text = (text or "").replace("\r", "").strip()
    if not text:
        return []

    c.setFont(font_name, font_size)

    def w(s: str) -> float:
        return c.stringWidth(s, font_name, font_size)

    tokens = text.split()
    lines: list[str] = []
    current = ""

    def push_line(line: str) -> None:
        if line.strip():
            lines.append(line.rstrip())

    def break_long_token(tok: str) -> list[str]:
        chunks: list[str] = []
        chunk = ""
        for ch in tok:
            test = chunk + ch
            if w(test) <= max_width or not chunk:
                chunk = test
            else:
                chunks.append(chunk)
                chunk = ch
        if chunk:
            chunks.append(chunk)
        return chunks

    for tok in tokens:
        if w(tok) > max_width:
            if current:
                push_line(current)
                current = ""
            for chunk in break_long_token(tok):
                push_line(chunk)
            continue

        candidate = tok if not current else f"{current} {tok}"
        if w(candidate) <= max_width:
            current = candidate
        else:
            push_line(current)
            current = tok

    if current:
        push_line(current)

    return lines


def _draw_wrapped_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    max_height: float,
    font_name: str,
    font_size: int,
    leading: int,
    valign: str = "top",
) -> None:
    text = (text or "").strip()
    if not text:
        return

    c.saveState()
    c.setFont(font_name, font_size)

    lines = _wrap_text_to_lines(c, text, max_width, font_name, font_size)
    if leading <= 0:
        leading = font_size + 2

    max_lines = int(max_height // leading) if max_height > 0 else len(lines)
    lines = lines[:max_lines]

    block_h = len(lines) * leading

    if valign == "bottom":
        start_y = y + 2 + (block_h - leading)
    elif valign == "middle":
        start_y = y + (max_height / 2.0) + (block_h / 2.0) - leading
    else:
        start_y = y + max_height - leading

    yy = start_y
    for ln in lines:
        c.drawString(x, yy, ln)
        yy -= leading

    c.restoreState()


def _draw_cell(
    c: canvas.Canvas,
    x: float,
    y: float,
    w: float,
    h: float,
    label: str,
    value: str,
    pad: float,
) -> None:
    c.saveState()
    c.setLineWidth(0.8)
    c.rect(x, y, w, h, stroke=1, fill=0)

    label_font = ("Helvetica-Bold", 8.8)
    value_font = ("Helvetica", 9.2)

    c.setFont(*label_font)
    label_y = y + h - (pad + 9)
    c.drawString(x + pad, label_y, str(label).strip())

    value_area_top = label_y - 6
    value_area_bottom = y + pad
    value_area_h = max(1.0, value_area_top - value_area_bottom)

    vf_name, vf_size = value_font
    leading = max(11, int(vf_size) + 2)

    lines = _wrap_text_to_lines(
        c=c,
        text=str(value or "").strip(),
        max_width=w - 2 * pad,
        font_name=vf_name,
        font_size=int(vf_size),
    )

    max_lines = int(value_area_h // leading) if leading > 0 else len(lines)
    lines = lines[:max_lines]

    block_h = len(lines) * leading
    start_y = value_area_bottom + (value_area_h / 2.0) + (block_h / 2.0) - leading

    c.setFont(vf_name, vf_size)
    yy = start_y
    for ln in lines:
        c.drawString(x + pad, yy, ln)
        yy -= leading

    c.restoreState()


def _draw_cell_wrapped(
    c: canvas.Canvas,
    x: float,
    y: float,
    w: float,
    h: float,
    label: str,
    value: str,
    pad: float,
    value_font: Tuple[str, int],
    valign: str = "top",
) -> None:
    c.saveState()
    c.setLineWidth(0.8)
    c.rect(x, y, w, h, stroke=1, fill=0)

    c.setFont("Helvetica-Bold", 9)
    label_y = y + h - (pad + 10)
    c.drawString(x + pad, label_y, str(label).strip())

    value_area_top = label_y - 6
    value_area_bottom = y + pad
    value_area_h = max(1.0, value_area_top - value_area_bottom)

    font_name, font_size = value_font
    leading = max(11, font_size + 2)

    _draw_wrapped_text(
        c=c,
        text=str(value or "").strip(),
        x=x + pad,
        y=value_area_bottom,
        max_width=w - 2 * pad,
        max_height=value_area_h,
        font_name=font_name,
        font_size=font_size,
        leading=leading,
        valign=valign,
    )

    c.restoreState()


def _draw_logo_in_box(
    c: canvas.Canvas,
    logo_path: str,
    x: float,
    y: float,
    w: float,
    h: float,
    pad: float,
) -> None:
    if not logo_path or not os.path.exists(logo_path):
        return

    try:
        with Image.open(logo_path) as im:
            lw, lh = im.size
    except Exception:
        return

    if lw <= 0 or lh <= 0:
        return

    box_w = max(1.0, w - 2 * pad)
    box_h = max(1.0, h - 2 * pad)

    logo_aspect = lw / lh
    box_aspect = box_w / box_h

    if logo_aspect > box_aspect:
        draw_w = box_w
        draw_h = box_w / logo_aspect
    else:
        draw_h = box_h
        draw_w = box_h * logo_aspect

    draw_x = x + (w - draw_w) / 2.0
    draw_y = y + (h - draw_h) / 2.0

    img = ImageReader(logo_path)
    c.drawImage(img, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")


# ---------------------------------------------------------------------
# ISO title block (uses title_block.drawing_title and title_block.comments if present)
# ---------------------------------------------------------------------

def _draw_iso_title_block(
    c: canvas.Canvas,
    page_w: float,
    page_h: float,
    title_block: TitleBlock,
    settings: ExportSettings,
    sheet_no: int,
    sheet_total: int,
) -> None:
    margin = settings.page_margin_pt
    tb_w = settings.title_block_width_pt

    tb_x = page_w - margin - tb_w
    tb_y = margin
    tb_h = page_h - 2 * margin

    c.saveState()
    c.setLineWidth(1.0)
    c.rect(tb_x, tb_y, tb_w, tb_h, stroke=1, fill=0)

    # Layout sizing
    sign_h = 95.0
    info_h = 90.0
    title_h = 150.0
    top_h = max(150.0, tb_h * 0.22)

    fixed = sign_h + info_h + title_h + top_h
    comments_h = max(140.0, tb_h - fixed)

    y0 = tb_y
    y1 = y0 + sign_h
    y2 = y1 + comments_h
    y3 = y2 + info_h
    y4 = y3 + title_h
    y5 = tb_y + tb_h

    for yy in (y1, y2, y3, y4):
        c.setLineWidth(0.9)
        c.line(tb_x, yy, tb_x + tb_w, yy)

    pad = 6.0

    # TOP: issuer/project/client + logo
    logo_col_w = tb_w * 0.42
    text_col_w = tb_w - logo_col_w
    logo_x = tb_x + text_col_w

    c.setLineWidth(0.8)
    c.line(logo_x, y4, logo_x, y5)

    cell_h = top_h / 3.0
    _draw_cell(c, tb_x, y5 - cell_h, text_col_w, cell_h, "ISSUER", getattr(title_block, "issuer_company", ""), pad)
    _draw_cell(c, tb_x, y5 - 2 * cell_h, text_col_w, cell_h, "PROJECT", getattr(title_block, "project", ""), pad)
    _draw_cell(c, tb_x, y5 - 3 * cell_h, text_col_w, cell_h, "CLIENT", getattr(title_block, "client", ""), pad)

    c.rect(logo_x, y4, logo_col_w, top_h, stroke=1, fill=0)
    _draw_logo_in_box(c, getattr(title_block, "logo_path", ""), logo_x, y4, logo_col_w, top_h, pad)

    # DRAWING TITLE (per-sheet)
    _draw_cell_wrapped(
        c=c,
        x=tb_x,
        y=y3,
        w=tb_w,
        h=title_h,
        label="DRAWING TITLE",
        value=getattr(title_block, "drawing_title", ""),
        pad=pad,
        value_font=("Helvetica-Bold", 12),
        valign="middle",
    )

    # INFO: DWG/REV then DATE (global)
    info_row_h = info_h / 2.0
    col_split = tb_x + tb_w * 0.62

    c.setLineWidth(0.8)
    c.line(tb_x, y2 + info_row_h, tb_x + tb_w, y2 + info_row_h)
    c.line(col_split, y2 + info_row_h, col_split, y3)

    _draw_cell(c, tb_x, y2 + info_row_h, col_split - tb_x, info_row_h, "DWG NO", getattr(title_block, "drawing_number", ""), pad)
    _draw_cell(c, col_split, y2 + info_row_h, tb_x + tb_w - col_split, info_row_h, "REV", getattr(title_block, "revision", ""), pad)
    _draw_cell(c, tb_x, y2, tb_w, info_row_h, "DATE", getattr(title_block, "date", ""), pad)

    # COMMENTS big box (per-sheet)
    _draw_cell_wrapped(
        c=c,
        x=tb_x,
        y=y1,
        w=tb_w,
        h=comments_h,
        label="COMMENTS / NOTES",
        value=getattr(title_block, "comments", ""),
        pad=pad,
        value_font=("Helvetica", 9),
        valign="top",
    )

    # SIGN (global)
    sign_row_h = sign_h / 2.0
    col_split2 = tb_x + tb_w * 0.62

    c.setLineWidth(0.8)
    c.line(tb_x, y0 + sign_row_h, tb_x + tb_w, y0 + sign_row_h)
    c.line(col_split2, y0, col_split2, y1)

    _draw_cell(c, tb_x, y0 + sign_row_h, col_split2 - tb_x, sign_row_h, "DRAWN", getattr(title_block, "drawn_by", ""), pad)
    _draw_cell(c, tb_x, y0, col_split2 - tb_x, sign_row_h, "CHECKED", getattr(title_block, "checked_by", ""), pad)
    _draw_cell(c, col_split2, y0 + sign_row_h, tb_x + tb_w - col_split2, sign_row_h, "APPROVED", getattr(title_block, "approved_by", ""), pad)
    _draw_cell(c, col_split2, y0, tb_x + tb_w - col_split2, sign_row_h, "SHEET", f"{sheet_no} of {sheet_total}", pad)

    c.restoreState()
