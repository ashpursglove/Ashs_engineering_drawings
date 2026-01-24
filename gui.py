
"""
gui.py

PyQt5 GUI for Engineering Drawing Maker (sheet plan + per-sheet overrides).

"""

from __future__ import annotations

import json
import os
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from PyQt5 import QtCore, QtGui, QtWidgets

from drawing_engine import export_sheet_plan_to_pdf
from models import ExportSettings, SheetPlanItem, TitleBlock
from templates import get_template_names


JOB_FILE_VERSION = 1
JOB_EXT = ".edmjob"






class MultiLineTextDelegate(QtWidgets.QStyledItemDelegate):
    """
    Delegate that uses QTextEdit for multi-line table cells.
    Ideal for comments / notes fields.
    """

    def createEditor(self, parent, option, index):
        editor = QtWidgets.QTextEdit(parent)
        editor.setWordWrapMode(QtGui.QTextOption.WordWrap)
        editor.setAcceptRichText(False)
        editor.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        return editor

    def setEditorData(self, editor, index):
        text = index.model().data(index, QtCore.Qt.EditRole) or ""
        editor.setPlainText(text)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), QtCore.Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)









class EngineeringDrawingMaker(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Ash's Engineering Drawing Maker")
        self.resize(1250, 800)

        self._files: list[str] = []
        self._sheet_plan: List[SheetPlanItem] = []

        # Job/save state
        self._job_path: Optional[str] = None
        self._dirty: bool = False
        self._loading: bool = False  # prevents "dirty" being set while populating UI

        self._build_ui()
        self._apply_dark_blue_style()
        self._connect_signals()
        self._install_toolbar()

        self._update_window_title()

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

        # >>> ADD THIS (Step 2) >>>
        comments_delegate = MultiLineTextDelegate(self.sheet_table)
        self.sheet_table.setItemDelegateForColumn(3, comments_delegate)

        # Make rows taller by default so comments are readable
        self.sheet_table.verticalHeader().setDefaultSectionSize(64)
        self.sheet_table.verticalHeader().setMinimumSectionSize(48)
        # <<< END ADD <<<

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

    def _install_toolbar(self) -> None:
        """
        A thin toolbar like a civilized CAD tool, not a ransom note.
        """
        tb = QtWidgets.QToolBar("Main")
        tb.setMovable(False)
        tb.setFloatable(False)
        tb.setIconSize(QtCore.QSize(16, 16))
        tb.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.addToolBar(QtCore.Qt.TopToolBarArea, tb)

        style = self.style()

        act_new = QtWidgets.QAction(style.standardIcon(QtWidgets.QStyle.SP_FileIcon), "New", self)
        act_new.setShortcut(QtGui.QKeySequence.New)
        act_new.triggered.connect(self._new_job)

        act_open = QtWidgets.QAction(style.standardIcon(QtWidgets.QStyle.SP_DialogOpenButton), "Open", self)
        act_open.setShortcut(QtGui.QKeySequence.Open)
        act_open.triggered.connect(self._open_job)

        act_save = QtWidgets.QAction(style.standardIcon(QtWidgets.QStyle.SP_DialogSaveButton), "Save", self)
        act_save.setShortcut(QtGui.QKeySequence.Save)
        act_save.triggered.connect(self._save_job)

        act_save_as = QtWidgets.QAction(style.standardIcon(QtWidgets.QStyle.SP_DialogSaveButton), "Save As", self)
        act_save_as.setShortcut(QtGui.QKeySequence("Ctrl+Shift+S"))
        act_save_as.triggered.connect(self._save_job_as)

        act_exit = QtWidgets.QAction(style.standardIcon(QtWidgets.QStyle.SP_DialogCloseButton), "Exit", self)
        act_exit.setShortcut(QtGui.QKeySequence.Quit)
        act_exit.triggered.connect(self.close)

        tb.addAction(act_new)
        tb.addAction(act_open)
        tb.addSeparator()
        tb.addAction(act_save)
        tb.addAction(act_save_as)
        tb.addSeparator()
        tb.addAction(act_exit)

        # Keep references if you later want enable/disable logic
        self._act_save = act_save
        self._act_save_as = act_save_as

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
        QToolBar {{
            background-color: {panel_bg};
            border-bottom: 1px solid {border_color};
            spacing: 6px;
            padding: 2px;
        }}
        QToolButton {{
            background-color: transparent;
            padding: 4px 8px;
            border-radius: 4px;
        }}
        QToolButton:hover {{
            background-color: {dark_bg_alt};
        }}
        QToolButton:pressed {{
            background-color: {highlight};
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

        # Dirty tracking for edits
        self.out_file.textChanged.connect(self._mark_dirty)
        self.cmb_template.currentIndexChanged.connect(self._mark_dirty)
        self.cmb_fit.currentIndexChanged.connect(self._mark_dirty)

        self.ed_issuer.textChanged.connect(self._mark_dirty)
        self.ed_logo.textChanged.connect(self._mark_dirty)

        self.ed_project.textChanged.connect(self._mark_dirty)
        self.ed_client.textChanged.connect(self._mark_dirty)
        self.ed_dwgno.textChanged.connect(self._mark_dirty)
        self.ed_rev.textChanged.connect(self._mark_dirty)
        self.ed_date.textChanged.connect(self._mark_dirty)
        self.ed_drawn.textChanged.connect(self._mark_dirty)
        self.ed_checked.textChanged.connect(self._mark_dirty)
        self.ed_approved.textChanged.connect(self._mark_dirty)

        self.sheet_table.itemChanged.connect(self._on_sheet_table_item_changed)

    # ------------------------------------------------------------------
    # Dirty / title
    # ------------------------------------------------------------------

    def _mark_dirty(self) -> None:
        if self._loading:
            return
        if not self._dirty:
            self._dirty = True
            self._update_window_title()

    def _set_clean(self) -> None:
        self._dirty = False
        self._update_window_title()

    def _update_window_title(self) -> None:
        name = "Untitled"
        if self._job_path:
            name = os.path.basename(self._job_path)
        star = " *" if self._dirty else ""
        self.setWindowTitle(f"Ash's Engineering Drawing Maker  [{name}{star}]")



    def _on_sheet_table_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        # Only mark dirty for user edits, not programmatic refresh
        if self._loading:
            return

        # If the user edited the Comments column, resize the row to fit contents
        if item and item.column() == 3:
            self.sheet_table.resizeRowToContents(item.row())

        self._mark_dirty()

    # ------------------------------------------------------------------
    # Job save/load (.edmjob)
    # ------------------------------------------------------------------

    def _maybe_save_changes(self) -> bool:
        """
        Returns True if it's ok to proceed (changes saved or discarded),
        False if user cancelled.
        """
        if not self._dirty:
            return True

        dlg = QtWidgets.QMessageBox(self)
        dlg.setIcon(QtWidgets.QMessageBox.Warning)
        dlg.setWindowTitle("Unsaved changes")
        dlg.setText("You have unsaved changes.")
        dlg.setInformativeText("Save them before continuing?")
        btn_save = dlg.addButton("Save", QtWidgets.QMessageBox.AcceptRole)
        btn_discard = dlg.addButton("Discard", QtWidgets.QMessageBox.DestructRole)
        btn_cancel = dlg.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
        dlg.setDefaultButton(btn_save)
        dlg.exec_()

        clicked = dlg.clickedButton()
        if clicked == btn_cancel:
            return False
        if clicked == btn_discard:
            return True
        if clicked == btn_save:
            return self._save_job()
        return False

    def _new_job(self) -> None:
        if not self._maybe_save_changes():
            return

        self._loading = True
        try:
            self._job_path = None
            self._files.clear()
            self._sheet_plan.clear()

            self.file_list.clear()
            self.sheet_table.setRowCount(0)

            self.out_file.clear()
            self.cmb_template.setCurrentText("A3_Landscape")
            self.cmb_fit.setCurrentText("fit")

            self.ed_issuer.clear()
            self.ed_logo.clear()

            self.ed_project.clear()
            self.ed_client.clear()
            self.ed_dwgno.clear()
            self.ed_rev.clear()
            self.ed_date.clear()
            self.ed_drawn.clear()
            self.ed_checked.clear()
            self.ed_approved.clear()

            self._set_status("Ready.")
        finally:
            self._loading = False

        self._set_clean()

    def _open_job(self) -> None:
        if not self._maybe_save_changes():
            return

        start_dir = os.path.dirname(self._job_path) if self._job_path else os.getcwd()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open job",
            start_dir,
            f"Engineering Drawing Maker Job (*{JOB_EXT});;JSON (*.json);;All files (*.*)",
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Open failed", f"Could not read job file:\n{exc}")
            return

        try:
            self._apply_job_dict(data)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Open failed", f"Job file is invalid:\n{exc}")
            return

        self._job_path = path
        self._set_clean()
        self._set_status(f"Loaded job: {os.path.basename(path)}")

    def _save_job(self) -> bool:
        """
        Saves to current job path; if none, falls back to Save As.
        Returns True on success.
        """
        if not self._job_path:
            return self._save_job_as()

        return self._write_job_file(self._job_path)

    def _save_job_as(self) -> bool:
        start_dir = os.path.dirname(self._job_path) if self._job_path else os.getcwd()
        suggested = os.path.join(start_dir, "drawing_job" + JOB_EXT)
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save job as",
            suggested,
            f"Engineering Drawing Maker Job (*{JOB_EXT});;All files (*.*)",
        )
        if not path:
            return False

        if not path.lower().endswith(JOB_EXT):
            path += JOB_EXT

        ok = self._write_job_file(path)
        if ok:
            self._job_path = path
            self._set_clean()
            self._set_status(f"Saved job: {os.path.basename(path)}")
        return ok

    def _write_job_file(self, path: str) -> bool:
        # Pull latest per-sheet overrides from UI table into model before saving
        self._sync_sheet_overrides_from_table()

        data = self._build_job_dict()

        out_dir = os.path.dirname(path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Save failed", f"Could not write job file:\n{exc}")
            return False

        self._set_clean()
        self._set_status(f"Saved: {os.path.basename(path)}")
        return True

    def _build_job_dict(self) -> dict:
        """
        Versioned, forward-friendly job schema.
        """
        # Save per-sheet overrides keyed by source_label (stable label used in table)
        overrides: Dict[str, Dict[str, str]] = {}
        for item in self._sheet_plan:
            overrides[item.source_label] = {
                "drawing_title": item.drawing_title,
                "comments": item.comments,
            }

        return {
            "version": JOB_FILE_VERSION,
            "files": list(self._files),
            "overrides_by_label": overrides,
            "ui": {
                "output_pdf": self.out_file.text().strip(),
                "template_name": self.cmb_template.currentText(),
                "fit_mode": self.cmb_fit.currentText(),
                "issuer_company": self.ed_issuer.text().strip(),
                "logo_path": self.ed_logo.text().strip(),
                "project": self.ed_project.text().strip(),
                "client": self.ed_client.text().strip(),
                "drawing_number": self.ed_dwgno.text().strip(),
                "revision": self.ed_rev.text().strip(),
                "date": self.ed_date.text().strip(),
                "drawn_by": self.ed_drawn.text().strip(),
                "checked_by": self.ed_checked.text().strip(),
                "approved_by": self.ed_approved.text().strip(),
            },
        }

    def _apply_job_dict(self, data: dict) -> None:
        """
        Loads a job dict (already parsed JSON).
        """
        if not isinstance(data, dict):
            raise ValueError("Job root must be a JSON object.")

        version = int(data.get("version", 0))
        if version != JOB_FILE_VERSION:
            raise ValueError(f"Unsupported job version: {version} (expected {JOB_FILE_VERSION})")

        files = data.get("files", [])
        if not isinstance(files, list):
            raise ValueError("Job 'files' must be a list.")

        overrides = data.get("overrides_by_label", {})
        if overrides is None:
            overrides = {}
        if not isinstance(overrides, dict):
            raise ValueError("Job 'overrides_by_label' must be a dict.")

        ui = data.get("ui", {})
        if ui is None:
            ui = {}
        if not isinstance(ui, dict):
            raise ValueError("Job 'ui' must be a dict.")

        self._loading = True
        try:
            self._files = [str(p) for p in files]
            self.file_list.clear()
            for p in self._files:
                self.file_list.addItem(p)

            # Rebuild plan from files (recounts PDF pages), then apply overrides by label
            self._rebuild_sheet_plan()

            for item in self._sheet_plan:
                ov = overrides.get(item.source_label)
                if isinstance(ov, dict):
                    item.drawing_title = str(ov.get("drawing_title", "") or "")
                    item.comments = str(ov.get("comments", "") or "")

            self._refresh_sheet_table()

            # UI fields
            self.out_file.setText(str(ui.get("output_pdf", "") or ""))
            self.cmb_template.setCurrentText(str(ui.get("template_name", "A3_Landscape") or "A3_Landscape"))
            self.cmb_fit.setCurrentText(str(ui.get("fit_mode", "fit") or "fit"))

            self.ed_issuer.setText(str(ui.get("issuer_company", "") or ""))
            self.ed_logo.setText(str(ui.get("logo_path", "") or ""))

            self.ed_project.setText(str(ui.get("project", "") or ""))
            self.ed_client.setText(str(ui.get("client", "") or ""))
            self.ed_dwgno.setText(str(ui.get("drawing_number", "") or ""))
            self.ed_rev.setText(str(ui.get("revision", "") or ""))
            self.ed_date.setText(str(ui.get("date", "") or ""))
            self.ed_drawn.setText(str(ui.get("drawn_by", "") or ""))
            self.ed_checked.setText(str(ui.get("checked_by", "") or ""))
            self.ed_approved.setText(str(ui.get("approved_by", "") or ""))
        finally:
            self._loading = False

        self._update_window_title()

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

        changed = False
        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self.file_list.addItem(p)
                changed = True

        if changed:
            self._rebuild_sheet_plan()
            self._set_status(f"{len(self._files)} file(s), {len(self._sheet_plan)} sheet(s).")
            self._mark_dirty()

    def _remove_selected(self) -> None:
        items = self.file_list.selectedItems()
        if not items:
            return

        changed = False
        for it in items:
            p = it.text()
            if p in self._files:
                self._files.remove(p)
                changed = True
            self.file_list.takeItem(self.file_list.row(it))

        if changed:
            self._rebuild_sheet_plan()
            self._set_status(f"{len(self._files)} file(s), {len(self._sheet_plan)} sheet(s).")
            self._mark_dirty()

    def _clear_files(self) -> None:
        if not self._files and not self._sheet_plan:
            return
        self._files.clear()
        self.file_list.clear()
        self._sheet_plan.clear()
        self.sheet_table.setRowCount(0)
        self._set_status("Cleared.")
        self._mark_dirty()

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
        self._loading = True
        try:
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
        finally:
            self._loading = False

    def _sync_sheet_overrides_from_table(self) -> None:
        """
        Pull per-sheet overrides from table back into _sheet_plan.
        Safe to call before save/export.
        """
        if not self._sheet_plan:
            return
        for row in range(self.sheet_table.rowCount()):
            if row >= len(self._sheet_plan):
                break
            title = (self.sheet_table.item(row, 2).text().strip() if self.sheet_table.item(row, 2) else "")
            comm = (self.sheet_table.item(row, 3).text().strip() if self.sheet_table.item(row, 3) else "")
            self._sheet_plan[row].drawing_title = title
            self._sheet_plan[row].comments = comm

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
            self._mark_dirty()

    def _choose_logo(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select logo", os.getcwd(), "Images (*.png *.jpg *.jpeg);;All files (*.*)"
        )
        if path:
            self.ed_logo.setText(path)
            self._mark_dirty()

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
        self._sync_sheet_overrides_from_table()

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

    # ------------------------------------------------------------------
    # Window close handling
    # ------------------------------------------------------------------

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        if self._maybe_save_changes():
            event.accept()
        else:
            event.ignore()




