"""Glasses Import Validator — desktop app.

Open an Excel file; it renders like a spreadsheet with every mistake the
validator would flag shaded (red = error, amber = warning), hover a cell for
the reason, and use the side panel to filter, jump between issues, refresh the
reference data, or export an annotated .xlsx.
"""
from __future__ import annotations
import sys
import pandas as pd

from PySide6.QtCore import Qt, QThread, Signal, QObject, QModelIndex, QSettings
from PySide6.QtGui import QAction, QColor, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QDockWidget, QWidget, QVBoxLayout,
    QHBoxLayout, QLabel, QPushButton, QCheckBox, QFileDialog, QMessageBox,
    QGroupBox, QGridLayout, QFrame, QProgressBar, QStyleFactory,
)

import remote
import updater
import validator_core as vc
from model import ValidationTableModel, IssueFilterProxy, ERROR_BG, WARNING_BG
from export import export_annotated
from version import __version__


def apply_theme(app, dark: bool):
    """Apply a light or dark Fusion palette to the whole app."""
    app.setStyle(QStyleFactory.create("Fusion"))
    pal = QPalette()
    if dark:
        c = QColor
        pal.setColor(QPalette.Window, c(45, 45, 48))
        pal.setColor(QPalette.WindowText, c(230, 230, 230))
        pal.setColor(QPalette.Base, c(30, 30, 32))
        pal.setColor(QPalette.AlternateBase, c(40, 40, 44))
        pal.setColor(QPalette.Text, c(230, 230, 230))
        pal.setColor(QPalette.Button, c(55, 55, 60))
        pal.setColor(QPalette.ButtonText, c(230, 230, 230))
        pal.setColor(QPalette.ToolTipBase, c(45, 45, 48))
        pal.setColor(QPalette.ToolTipText, c(230, 230, 230))
        pal.setColor(QPalette.Highlight, c(38, 110, 190))
        pal.setColor(QPalette.HighlightedText, c(255, 255, 255))
        pal.setColor(QPalette.Disabled, QPalette.Text, c(130, 130, 130))
    else:
        pal = app.style().standardPalette()
    app.setPalette(pal)


# ---------- generic background worker ----------
class Worker(QThread):
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            self.done.emit(self._fn())
        except Exception as e:  # noqa
            self.failed.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Glasses Import Validator (Desktop) v{__version__}")
        self.resize(1400, 850)

        self.master_df = None
        self.user_df = None
        self.result = None
        self.model = None
        self.proxy = IssueFilterProxy()
        self._issue_cells = []
        self._issue_pos = -1
        self._worker = None
        self._dirty = False
        self._current_name = None
        self.settings = QSettings("Alensa", "GlassesValidator")

        # ---- central table ----
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(False)
        self.table.setAlternatingRowColors(True)
        self.setCentralWidget(self.table)

        self._build_toolbar()
        self._build_side_panel()
        self._build_statusbar()

        # Apply saved theme
        dark = self.settings.value("dark_mode", False, type=bool)
        self.act_dark.setChecked(dark)
        apply_theme(QApplication.instance(), dark)

        self._load_reference_data()
        self._check_updates_async()

    # ---------- UI construction ----------
    def _build_toolbar(self):
        tb = self.addToolBar("Main")
        tb.setMovable(False)
        act_open = QAction("📂 Open file", self); act_open.triggered.connect(self.open_file); tb.addAction(act_open)
        act_reval = QAction("🔁 Re-validate", self); act_reval.triggered.connect(self.run_validation); tb.addAction(act_reval)
        act_fix = QAction("🧹 Fix all spacing", self); act_fix.triggered.connect(self.fix_spacing); tb.addAction(act_fix)
        act_save = QAction("💾 Save changes", self); act_save.triggered.connect(self.save_file); tb.addAction(act_save)
        tb.addSeparator()
        act_refresh = QAction("☁️ Refresh data", self); act_refresh.triggered.connect(self.refresh_data); tb.addAction(act_refresh)
        act_export = QAction("💾 Export annotated", self); act_export.triggered.connect(self.export_file); tb.addAction(act_export)
        tb.addSeparator()
        act_upd = QAction("⬆️ Check updates", self); act_upd.triggered.connect(lambda: self._check_updates_async(manual=True)); tb.addAction(act_upd)
        self.act_dark = QAction("🌙 Dark mode", self); self.act_dark.setCheckable(True)
        self.act_dark.toggled.connect(self._toggle_dark); tb.addAction(self.act_dark)

    def _build_statusbar(self):
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(180)
        self.progress.setRange(0, 0)   # busy/indeterminate
        self.progress.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress)

    def _set_busy(self, on: bool, message: str = ""):
        self.progress.setVisible(on)
        if message:
            self.statusBar().showMessage(message, 0 if on else 8000)

    def _toggle_dark(self, on: bool):
        apply_theme(QApplication.instance(), on)
        self.settings.setValue("dark_mode", on)

    def _build_side_panel(self):
        dock = QDockWidget("Control panel", self)
        dock.setFeatures(QDockWidget.NoDockWidgetFeatures)
        panel = QWidget(); lay = QVBoxLayout(panel)

        # Reference-data status (prominent)
        self.lbl_data_status = QLabel("⏳ Loading reference data…")
        self.lbl_data_status.setWordWrap(True)
        self.lbl_data_status.setStyleSheet(
            "padding:8px; border-radius:6px; background:#8a6d000f; "
            "color:#a06f00; font-weight:bold;")
        lay.addWidget(self.lbl_data_status)

        # Selected-cell details
        self.grp_cell = QGroupBox("Selected cell")
        cl = QVBoxLayout(self.grp_cell)
        self.lbl_cell = QLabel("Click a highlighted cell to see its issue.")
        self.lbl_cell.setWordWrap(True)
        self.lbl_cell.setTextFormat(Qt.RichText)
        cl.addWidget(self.lbl_cell)
        lay.addWidget(self.grp_cell)

        # Summary
        self.grp_summary = QGroupBox("Summary")
        g = QGridLayout(self.grp_summary)
        self.lbl_rows = QLabel("—"); self.lbl_total = QLabel("—")
        self.lbl_empty = QLabel("—"); self.lbl_invalid = QLabel("—")
        self.lbl_case = QLabel("—"); self.lbl_ws = QLabel("—"); self.lbl_meta = QLabel("—")
        rows = [("Rows", self.lbl_rows), ("Total issues", self.lbl_total),
                ("Empty required", self.lbl_empty), ("Invalid values", self.lbl_invalid),
                ("Case mismatch", self.lbl_case), ("Whitespace", self.lbl_ws),
                ("Meta format", self.lbl_meta)]
        for i, (name, lbl) in enumerate(rows):
            g.addWidget(QLabel(name + ":"), i, 0); g.addWidget(lbl, i, 1)
        lay.addWidget(self.grp_summary)

        # Filters
        grp_f = QGroupBox("View")
        fl = QVBoxLayout(grp_f)
        self.chk_only = QCheckBox("Show only rows with issues")
        self.chk_only.toggled.connect(self.proxy.set_only_issues)
        fl.addWidget(self.chk_only)
        btn_next = QPushButton("⏭ Jump to next issue"); btn_next.clicked.connect(self.jump_next_issue)
        fl.addWidget(btn_next)
        lay.addWidget(grp_f)

        # Legend
        grp_l = QGroupBox("Legend")
        ll = QVBoxLayout(grp_l)
        ll.addWidget(self._legend_row(ERROR_BG, "Error (empty required, invalid value, meta format)"))
        ll.addWidget(self._legend_row(WARNING_BG, "Warning (whitespace, case mismatch)"))
        lay.addWidget(grp_l)

        # Mapping info
        self.grp_map = QGroupBox("Column mapping")
        ml = QVBoxLayout(self.grp_map)
        self.lbl_map = QLabel("—"); self.lbl_map.setWordWrap(True)
        ml.addWidget(self.lbl_map)
        lay.addWidget(self.grp_map)

        lay.addStretch(1)
        dock.setWidget(panel)
        dock.setMinimumWidth(300)
        self.addDockWidget(Qt.RightDockWidgetArea, dock)

    def _legend_row(self, color: QColor, text: str):
        w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0, 0, 0, 0)
        sw = QFrame(); sw.setFixedSize(18, 18)
        sw.setStyleSheet(f"background-color: {color.name()}; border: 1px solid #888;")
        h.addWidget(sw); lbl = QLabel(text); lbl.setWordWrap(True); h.addWidget(lbl, 1)
        return w

    # ---------- reference data ----------
    def _set_data_status(self, text, state):
        colors = {
            "loading": ("#a06f00", "#8a6d000f"),
            "ready": ("#1a7f37", "#1a7f370f"),
            "error": ("#b30000", "#b300000f"),
        }
        fg, bg = colors.get(state, colors["loading"])
        self.lbl_data_status.setText(text)
        self.lbl_data_status.setStyleSheet(
            f"padding:8px; border-radius:6px; background:{bg}; color:{fg}; font-weight:bold;")

    def _load_reference_data(self, force=False):
        self._set_busy(True, "Loading reference data from GitHub…")
        self._set_data_status("⏳ Loading reference data…", "loading")
        self._worker = Worker(lambda: remote.load_master(force=force))
        self._worker.done.connect(self._on_master_loaded)
        self._worker.failed.connect(self._on_data_error)
        self._worker.start()

    def _on_master_loaded(self, df):
        self.master_df = df
        self._set_busy(False)
        status = remote.LAST_STATUS.get("master_clean.xlsx", "downloaded")
        note = {
            "up-to-date": "cached · up to date",
            "updated": "updated from repo",
            "downloaded": "downloaded",
            "offline-cache": "offline · using cached copy",
        }.get(status, status)
        self._set_data_status(f"✅ Reference data ready\n{len(df):,} master rows ({note})", "ready")
        self.statusBar().showMessage("Reference data ready — open a file to validate.", 8000)
        if self.user_df is not None:
            self.run_validation()

    def _on_data_error(self, e):
        self._set_busy(False)
        self._set_data_status("❌ Could not load reference data\n(check your connection)", "error")
        self._error("Could not load master data", e)

    def refresh_data(self):
        remote.force_refresh()
        self._load_reference_data(force=True)

    # ---------- file open + validation ----------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, dtype=str, header=0).reset_index(drop=True)
            # Normalise headers the same way the validator does, so the grid
            # column names match the keys used for cell issues (headers in these
            # files often contain line breaks / double spaces).
            self.user_df = vc.clean_headers(df)
        except Exception as e:
            self._error("Could not read file", str(e)); return
        self._current_path = path
        self._current_name = path.replace("\\", "/").split("/")[-1]
        self._dirty = False
        self.setWindowTitle(self._base_title())
        self.run_validation()

    def run_validation(self):
        if self.user_df is None:
            QMessageBox.information(self, "No file", "Open an Excel file first."); return
        if self.master_df is None:
            QMessageBox.information(self, "Please wait", "Reference data is still loading."); return
        self._set_busy(True, "Validating…")
        udf, mdf = self.user_df, self.master_df
        self._worker = Worker(lambda: vc.validate(udf, mdf))
        self._worker.done.connect(self._on_validated)
        self._worker.failed.connect(lambda e: (self._set_busy(False), self._error("Validation failed", e)))
        self._worker.start()

    def _on_validated(self, result):
        self._set_busy(False)
        self.result = result
        self.model = ValidationTableModel(self.user_df, result["cell_issues"])
        # The model edits this DataFrame in place, so keep the same object.
        self.user_df = self.model.dataframe()
        self.model.edited.connect(self._on_model_edited)
        self.proxy.setSourceModel(self.model)
        self.table.selectionModel().currentChanged.connect(self._on_cell_selected)
        self.proxy.set_only_issues(self.chk_only.isChecked())
        self.table.resizeColumnsToContents()
        self._issue_cells = sorted(result["cell_issues"].keys())
        self._issue_pos = -1
        self._update_summary(result)
        c = result["counts"]
        self.statusBar().showMessage(f"Done — {c['total']} issue(s) found.", 8000)

    def _update_summary(self, result):
        c = result["counts"]
        self.lbl_rows.setText(str(len(self.user_df)))
        self.lbl_total.setText(str(c["total"]))
        self.lbl_empty.setText(str(c["empty"]))
        self.lbl_invalid.setText(str(c["invalid"]))
        self.lbl_case.setText(str(c["case_mismatch"]))
        self.lbl_ws.setText(str(c["whitespace"]))
        self.lbl_meta.setText(str(c["meta_format"]))
        m = result["mapping"]
        self.lbl_map.setText(
            f"Mapped {len(m['active'])}/{len(vc.IDEAL_PAIRS)} columns · "
            f"required {len(m['required_found'])}/{len(vc.REQUIRED_COLUMN_KEYWORDS)} · "
            f"sunglasses {len(m['sunglasses_found'])}/{len(vc.SUNGLASSES_REQUIRED_KEYWORDS)}"
            + (f"\n⚠️ {len(m['unmapped'])} unmapped pair(s)." if m["unmapped"] else "")
        )

    def _on_cell_selected(self, current, previous):
        if self.model is None or not current.isValid():
            self.lbl_cell.setText("Click a highlighted cell to see its issue.")
            return
        src = self.proxy.mapToSource(current)
        cols = list(self.model.dataframe().columns)
        col_name = cols[src.column()]
        value = self.model.dataframe().iat[src.row(), src.column()]
        value = "" if value is None or str(value) == "nan" else str(value)
        issues = self.model.issues_at(src.row(), col_name)

        html = [f"<b>Row {src.row() + 2}</b> · <b>{col_name}</b>",
                f"Value: <code>{(value or '(empty)')}</code>"]
        if not issues:
            html.append("<span style='color:#1a7f37'>✓ No issues on this cell.</span>")
        else:
            for i in issues:
                sev = "#b30000" if i["severity"] == "error" else "#a06f00"
                html.append(f"<div style='margin-top:6px'>"
                            f"<span style='color:{sev}'><b>Issue:</b> {i['message']}</span>")
                if i.get("expected"):
                    html.append(f"<b>Expected:</b> {i['expected']}")
                html.append("</div>")
        self.lbl_cell.setText("<br>".join(html))

    def jump_next_issue(self):
        if not self._issue_cells or self.model is None:
            return
        self._issue_pos = (self._issue_pos + 1) % len(self._issue_cells)
        row, col_name = self._issue_cells[self._issue_pos]
        try:
            col = list(self.model.dataframe().columns).index(col_name)
        except ValueError:
            return
        src = self.model.index(row, col)
        pidx = self.proxy.mapFromSource(src)
        if pidx.isValid():
            self.table.setCurrentIndex(pidx)
            self.table.scrollTo(pidx)

    def fix_spacing(self):
        """One-click cleanup of leading/trailing/double spaces, NBSP and
        spaces around the '|' separator, across the whole sheet."""
        if self.user_df is None:
            QMessageBox.information(self, "No file", "Open an Excel file first."); return
        fixed_df, changes = vc.fix_spacing(self.user_df)
        if not changes:
            QMessageBox.information(self, "Nothing to fix", "No spacing issues found.")
            return
        cols = sorted({c for _r, c, _b, _a in changes})
        msg = (f"Fix {len(changes)} cell(s) across {len(cols)} column(s)?\n\n"
               "This trims leading/trailing spaces, collapses double spaces,\n"
               "replaces non-breaking spaces, and tidies spaces around '|'.")
        if QMessageBox.question(self, "Fix all spacing", msg) != QMessageBox.Yes:
            return
        self.user_df = fixed_df
        self._dirty = True
        self.statusBar().showMessage(f"Fixed spacing in {len(changes)} cell(s) — re-validating…", 5000)
        self.run_validation()

    def _on_model_edited(self):
        self._dirty = True
        self.setWindowTitle(self._base_title() + " •")

    def _base_title(self):
        name = f" — {self._current_name}" if getattr(self, "_current_name", None) else ""
        return f"Glasses Import Validator (Desktop) v{__version__}{name}"

    def save_file(self):
        """Write the edited table back to .xlsx (plain values, no highlighting)."""
        if self.user_df is None:
            QMessageBox.information(self, "No file", "Open an Excel file first."); return
        default = getattr(self, "_current_path", "edited.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "Save edited file", default, "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.user_df.to_excel(path, index=False)
            self._dirty = False
            self.setWindowTitle(self._base_title())
            self.statusBar().showMessage(f"Saved → {path}", 8000)
        except Exception as e:
            self._error("Save failed", str(e))

    def export_file(self):
        if self.result is None:
            QMessageBox.information(self, "Nothing to export", "Validate a file first."); return
        path, _ = QFileDialog.getSaveFileName(self, "Save annotated Excel", "validated.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            export_annotated(self.user_df, self.result["cell_issues"], path)
            self.statusBar().showMessage(f"Exported → {path}", 8000)
        except Exception as e:
            self._error("Export failed", str(e))

    # ---------- updates ----------
    def _check_updates_async(self, manual=False):
        self._upd_worker = Worker(updater.check_latest)
        self._upd_worker.done.connect(lambda info: self._on_update_info(info, manual))
        if manual:
            self._upd_worker.failed.connect(lambda e: self._error("Update check failed", e))
        self._upd_worker.start()

    def _on_update_info(self, info, manual):
        if not info:
            if manual:
                QMessageBox.information(self, "Updates", "Could not reach the update server.")
            return
        if not info.get("newer"):
            if manual:
                QMessageBox.information(self, "Updates", f"You're on the latest version (v{__version__}).")
            return
        msg = f"A new version ({info['tag']}) is available.\n\n{info.get('notes','')[:500]}\n\nUpdate now?"
        if QMessageBox.question(self, "Update available", msg) == QMessageBox.Yes:
            if info.get("exe_url") and updater.download_and_apply(info["exe_url"]):
                QMessageBox.information(self, "Updating", "The app will now close and relaunch.")
                QApplication.quit()
            else:
                QMessageBox.warning(self, "Update", "Automatic update only works in the packaged .exe.")

    # ---------- misc ----------
    def closeEvent(self, event):
        if self._dirty:
            r = QMessageBox.question(
                self, "Unsaved changes",
                "You have unsaved edits. Save before closing?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            if r == QMessageBox.Cancel:
                event.ignore(); return
            if r == QMessageBox.Save:
                self.save_file()
        event.accept()

    def _error(self, title, detail):
        self.statusBar().showMessage(title, 6000)
        QMessageBox.critical(self, title, str(detail))


def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
