

"""
gui.py

PyQt5 GUI for Engineering Drawing Maker (sheet plan + per-sheet overrides).
"""

from __future__ import annotations

import os
from typing import List

import fitz  # PyMuPDF
from PyQt5 import QtCore, QtWidgets

from drawing_engine import export_sheet_plan_to_pdf
from models import ExportSettings, SheetPlanItem, TitleBlock
from templates import get_template_names


class EngineeringDrawingMaker(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Ash's Engineering Drawing Maker")
        self.resize(1250, 800)

        self._files: list[str] = []
        self._sheet_plan: List[SheetPlanItem] = []

        self._build_ui()
        self._apply_dark_blue_style()
        self._connect_signals()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        root.addWidget(split)

        # LEFT: files + sheet overrides table
        left = QtWidgets.QWidget()
        left_l = QtWidgets.QVBoxLayout(left)
        left_l.setContentsMargins(0, 0, 0, 0)
        left_l.setSpacing(8)

        title = QtWidgets.QLabel("Inputs and Sheets")
        f = title.font()
        f.setBold(True)
        f.setPointSize(11)
        title.setFont(f)
        left_l.addWidget(title)

        self.file_list = QtWidgets.QListWidget()
        self.file_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        left_l.addWidget(self.file_list, 2)

        btn_row = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Add files")
        self.btn_remove = QtWidgets.QPushButton("Remove selected")
        self.btn_clear = QtWidgets.QPushButton("Clear")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        left_l.addLayout(btn_row)

        sheet_lbl = QtWidgets.QLabel("Per-sheet overrides (Drawing Title + Comments)")
        sf = sheet_lbl.font()
        sf.setBold(True)
        sheet_lbl.setFont(sf)
        left_l.addWidget(sheet_lbl)

        self.sheet_table = QtWidgets.QTableWidget(0, 4)
        self.sheet_table.setHorizontalHeaderLabels(["Sheet", "Source", "Drawing Title", "Comments"])
        self.sheet_table.horizontalHeader().setStretchLastSection(True)
        self.sheet_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.sheet_table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.sheet_table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        self.sheet_table.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.sheet_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.sheet_table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed)
        left_l.addWidget(self.sheet_table, 3)

        # RIGHT: global settings
        right = QtWidgets.QWidget()
        right_l = QtWidgets.QVBoxLayout(right)
        right_l.setContentsMargins(0, 0, 0, 0)
        right_l.setSpacing(10)

        out_g = QtWidgets.QGroupBox("Output")
        out_f = QtWidgets.QFormLayout(out_g)
        out_f.setLabelAlignment(QtCore.Qt.AlignRight)

        self.out_file = QtWidgets.QLineEdit()
        self.out_file.setPlaceholderText("Where should the combined PDF pack be saved?")
        self.btn_out_file = QtWidgets.QPushButton("Browse")

        out_row = QtWidgets.QHBoxLayout()
        out_row.addWidget(self.out_file, 1)
        out_row.addWidget(self.btn_out_file)
        out_f.addRow("Output PDF:", out_row)

        self.cmb_template = QtWidgets.QComboBox()
        self.cmb_template.addItems(get_template_names())
        self.cmb_template.setCurrentText("A3_Landscape")
        out_f.addRow("Template:", self.cmb_template)

        self.cmb_fit = QtWidgets.QComboBox()
        self.cmb_fit.addItems(["fit", "fill"])
        out_f.addRow("Image fit mode:", self.cmb_fit)

        right_l.addWidget(out_g)

        issuer_g = QtWidgets.QGroupBox("Issuer")
        issuer_f = QtWidgets.QFormLayout(issuer_g)
        issuer_f.setLabelAlignment(QtCore.Qt.AlignRight)

        self.ed_issuer = QtWidgets.QLineEdit()
        issuer_f.addRow("Issuer company:", self.ed_issuer)

        self.ed_logo = QtWidgets.QLineEdit()
        self.ed_logo.setPlaceholderText("Optional logo (PNG/JPG)")
        self.btn_logo = QtWidgets.QPushButton("Browse")

        logo_row = QtWidgets.QHBoxLayout()
        logo_row.addWidget(self.ed_logo, 1)
        logo_row.addWidget(self.btn_logo)
        issuer_f.addRow("Logo file:", logo_row)

        right_l.addWidget(issuer_g)

        tb_g = QtWidgets.QGroupBox("Global Title Block Fields (apply to all sheets)")
        tb_f = QtWidgets.QFormLayout(tb_g)
        tb_f.setLabelAlignment(QtCore.Qt.AlignRight)

        self.ed_project = QtWidgets.QLineEdit()
        self.ed_client = QtWidgets.QLineEdit()
        self.ed_dwgno = QtWidgets.QLineEdit()
        self.ed_rev = QtWidgets.QLineEdit()
        self.ed_date = QtWidgets.QLineEdit()
        self.ed_drawn = QtWidgets.QLineEdit()
        self.ed_checked = QtWidgets.QLineEdit()
        self.ed_approved = QtWidgets.QLineEdit()

        tb_f.addRow("Project:", self.ed_project)
        tb_f.addRow("Client:", self.ed_client)
        tb_f.addRow("Drawing number:", self.ed_dwgno)
        tb_f.addRow("Revision:", self.ed_rev)
        tb_f.addRow("Date:", self.ed_date)
        tb_f.addRow("Drawn by:", self.ed_drawn)
        tb_f.addRow("Checked by:", self.ed_checked)
        tb_f.addRow("Approved by:", self.ed_approved)

        right_l.addWidget(tb_g)

        bottom = QtWidgets.QHBoxLayout()
        self.btn_export = QtWidgets.QPushButton("Export PDF Pack")
        self.lbl_status = QtWidgets.QLabel("Ready.")
        bottom.addWidget(self.btn_export)
        bottom.addWidget(self.lbl_status, 1)
        right_l.addLayout(bottom)

        split.addWidget(left)
        split.addWidget(right)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)

    def _apply_dark_blue_style(self) -> None:
        dark_bg = "#0b1f3b"
        dark_bg_alt = "#12284a"
        panel_bg = "#16243a"
        text_color = "#f0f4ff"
        accent = "#ff8c00"
        border_color = "#2b3f5f"
        highlight = "#2854a0"

        qss = f"""
        QWidget {{
            background-color: {dark_bg};
            color: {text_color};
            selection-background-color: {highlight};
            selection-color: {text_color};
        }}
        QGroupBox {{
            background-color: {panel_bg};
            border: 1px solid {border_color};
            border-radius: 4px;
            margin-top: 6px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 2px 6px;
            color: {accent};
            font-weight: bold;
        }}
        QListWidget, QTableWidget {{
            background-color: {dark_bg_alt};
            border: 1px solid {border_color};
        }}
        QHeaderView::section {{
            background-color: {panel_bg};
            color: {text_color};
            border: 1px solid {border_color};
            padding: 4px;
        }}
        QLineEdit, QComboBox {{
            background-color: {dark_bg_alt};
            border: 1px solid {border_color};
            border-radius: 3px;
            padding: 4px 6px;
        }}
        QPushButton {{
            background-color: {accent};
            color: #000000;
            border-radius: 4px;
            padding: 6px 12px;
            font-weight: bold;
        }}
        QPushButton:hover {{ background-color: #ff9e26; }}
        QPushButton:pressed {{ background-color: #e67a00; }}
        """
        self.setStyleSheet(qss)

    def _connect_signals(self) -> None:
        self.btn_add.clicked.connect(self._add_files)
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_clear.clicked.connect(self._clear_files)
        self.btn_out_file.clicked.connect(self._choose_out_file)
        self.btn_logo.clicked.connect(self._choose_logo)
        self.btn_export.clicked.connect(self._export)

    # ------------------------------------------------------------------
    # File selection + sheet plan building
    # ------------------------------------------------------------------

    def _add_files(self) -> None:
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Select files",
            os.getcwd(),
            "Supported (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.pdf);;All files (*.*)",
        )
        if not paths:
            return

        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self.file_list.addItem(p)

        self._rebuild_sheet_plan()
        self._set_status(f"{len(self._files)} file(s), {len(self._sheet_plan)} sheet(s).")

    def _remove_selected(self) -> None:
        items = self.file_list.selectedItems()
        if not items:
            return
        for it in items:
            p = it.text()
            if p in self._files:
                self._files.remove(p)
            self.file_list.takeItem(self.file_list.row(it))

        self._rebuild_sheet_plan()
        self._set_status(f"{len(self._files)} file(s), {len(self._sheet_plan)} sheet(s).")

    def _clear_files(self) -> None:
        self._files.clear()
        self.file_list.clear()
        self._sheet_plan.clear()
        self.sheet_table.setRowCount(0)
        self._set_status("Cleared.")

    def _rebuild_sheet_plan(self) -> None:
        """
        Build sheet plan with 1 row per output sheet.
        For PDFs we add one sheet per page (no rendering here, just page counts).
        """
        # Preserve existing overrides by source_label if possible
        old_by_label = {i.source_label: i for i in self._sheet_plan}

        plan: List[SheetPlanItem] = []
        for p in self._files:
            ext = os.path.splitext(p)[1].lower()
            base = os.path.basename(p)

            if ext == ".pdf":
                try:
                    doc = fitz.open(p)
                    pc = doc.page_count
                    doc.close()
                except Exception:
                    pc = 0

                for i in range(pc):
                    label = f"{base} - Page {i+1}"
                    prev = old_by_label.get(label)
                    plan.append(
                        SheetPlanItem(
                            kind="pdf",
                            source_path=p,
                            source_label=label,
                            pdf_page_index=i,
                            drawing_title=(prev.drawing_title if prev else ""),
                            comments=(prev.comments if prev else ""),
                        )
                    )
            else:
                label = base
                prev = old_by_label.get(label)
                plan.append(
                    SheetPlanItem(
                        kind="image",
                        source_path=p,
                        source_label=label,
                        pdf_page_index=None,
                        drawing_title=(prev.drawing_title if prev else ""),
                        comments=(prev.comments if prev else ""),
                    )
                )

        self._sheet_plan = plan
        self._refresh_sheet_table()

    def _refresh_sheet_table(self) -> None:
        self.sheet_table.setRowCount(0)

        for idx, item in enumerate(self._sheet_plan, start=1):
            row = self.sheet_table.rowCount()
            self.sheet_table.insertRow(row)

            # Sheet #
            it_sheet = QtWidgets.QTableWidgetItem(str(idx))
            it_sheet.setFlags(it_sheet.flags() & ~QtCore.Qt.ItemIsEditable)
            self.sheet_table.setItem(row, 0, it_sheet)

            # Source
            it_src = QtWidgets.QTableWidgetItem(item.source_label)
            it_src.setFlags(it_src.flags() & ~QtCore.Qt.ItemIsEditable)
            self.sheet_table.setItem(row, 1, it_src)

            # Drawing Title (editable)
            it_title = QtWidgets.QTableWidgetItem(item.drawing_title)
            self.sheet_table.setItem(row, 2, it_title)

            # Comments (editable)
            it_comm = QtWidgets.QTableWidgetItem(item.comments)
            self.sheet_table.setItem(row, 3, it_comm)

    # ------------------------------------------------------------------
    # Output selection / logo
    # ------------------------------------------------------------------

    def _choose_out_file(self) -> None:
        suggested = os.path.join(os.getcwd(), "drawing_pack.pdf")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save PDF pack", suggested, "PDF (*.pdf)")
        if path:
            if not path.lower().endswith(".pdf"):
                path += ".pdf"
            self.out_file.setText(path)

    def _choose_logo(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select logo", os.getcwd(), "Images (*.png *.jpg *.jpeg);;All files (*.*)"
        )
        if path:
            self.ed_logo.setText(path)

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def _export(self) -> None:
        if not self._sheet_plan:
            QtWidgets.QMessageBox.warning(self, "Nothing to export", "Add some files first.")
            return

        out_pdf = self.out_file.text().strip()
        if not out_pdf:
            QtWidgets.QMessageBox.warning(self, "No output file", "Choose where to save the PDF pack.")
            return

        if not out_pdf.lower().endswith(".pdf"):
            out_pdf += ".pdf"

        out_dir = os.path.dirname(out_pdf)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Pull per-sheet overrides from table back into plan
        for row in range(self.sheet_table.rowCount()):
            title = (self.sheet_table.item(row, 2).text().strip() if self.sheet_table.item(row, 2) else "")
            comm = (self.sheet_table.item(row, 3).text().strip() if self.sheet_table.item(row, 3) else "")
            self._sheet_plan[row].drawing_title = title
            self._sheet_plan[row].comments = comm

        global_tb = TitleBlock(
            issuer_company=self.ed_issuer.text().strip(),
            logo_path=self.ed_logo.text().strip(),
            project=self.ed_project.text().strip(),
            client=self.ed_client.text().strip(),
            drawing_number=self.ed_dwgno.text().strip(),
            revision=self.ed_rev.text().strip(),
            date=self.ed_date.text().strip(),
            drawn_by=self.ed_drawn.text().strip(),
            checked_by=self.ed_checked.text().strip(),
            approved_by=self.ed_approved.text().strip(),
        )

        settings = ExportSettings(
            template_name=self.cmb_template.currentText(),
            fit_mode=self.cmb_fit.currentText(),
        )

        self._set_status("Exporting...")
        QtWidgets.QApplication.processEvents()

        try:
            export_sheet_plan_to_pdf(
                sheet_plan=self._sheet_plan,
                output_pdf_path=out_pdf,
                global_tb=global_tb,
                settings=settings,
            )
        except Exception as exc:
            self._set_status("Failed.")
            QtWidgets.QMessageBox.critical(self, "Export failed", str(exc))
            return

        self._set_status("Done.")
        QtWidgets.QMessageBox.information(self, "Export completed", f"Created drawing pack:\n{out_pdf}")

    def _set_status(self, msg: str) -> None:
        self.lbl_status.setText(msg)








