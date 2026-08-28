"""Glasses Import Validator — desktop app.

Open an Excel file; it renders like a spreadsheet with every mistake the
validator would flag shaded (red = error, amber = warning), hover a cell for
the reason, and use the side panel to filter, jump between issues, refresh the
reference data, or export an annotated .xlsx.
"""
from __future__ import annotations
import sys
import pandas as pd

from PySide6.QtCore import Qt, QThread, Signal, QObject, QModelIndex
from PySide6.QtGui import QAction, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QDockWidget, QWidget, QVBoxLayout,
    QHBoxLayout, QLabel, QPushButton, QCheckBox, QFileDialog, QMessageBox,
    QGroupBox, QGridLayout, QFrame, QProgressDialog,
)

import remote
import updater
import validator_core as vc
from model import ValidationTableModel, IssueFilterProxy, ERROR_BG, WARNING_BG
from export import export_annotated
from version import __version__


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

        # ---- central table ----
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(False)
        self.table.setAlternatingRowColors(True)
        self.setCentralWidget(self.table)

        self._build_toolbar()
        self._build_side_panel()
        self.statusBar().showMessage("Loading reference data…")

        self._load_reference_data()
        self._check_updates_async()

    # ---------- UI construction ----------
    def _build_toolbar(self):
        tb = self.addToolBar("Main")
        tb.setMovable(False)
        act_open = QAction("📂 Open file", self); act_open.triggered.connect(self.open_file); tb.addAction(act_open)
        act_reval = QAction("🔁 Re-validate", self); act_reval.triggered.connect(self.run_validation); tb.addAction(act_reval)
        tb.addSeparator()
        act_refresh = QAction("☁️ Refresh data", self); act_refresh.triggered.connect(self.refresh_data); tb.addAction(act_refresh)
        act_export = QAction("💾 Export annotated", self); act_export.triggered.connect(self.export_file); tb.addAction(act_export)
        tb.addSeparator()
        act_upd = QAction("⬆️ Check updates", self); act_upd.triggered.connect(lambda: self._check_updates_async(manual=True)); tb.addAction(act_upd)

    def _build_side_panel(self):
        dock = QDockWidget("Control panel", self)
        dock.setFeatures(QDockWidget.NoDockWidgetFeatures)
        panel = QWidget(); lay = QVBoxLayout(panel)

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
    def _load_reference_data(self, force=False):
        self.statusBar().showMessage("Loading reference data from GitHub…")
        self._worker = Worker(lambda: remote.load_master(force=force))
        self._worker.done.connect(self._on_master_loaded)
        self._worker.failed.connect(lambda e: self._error("Could not load master data", e))
        self._worker.start()

    def _on_master_loaded(self, df):
        self.master_df = df
        self.statusBar().showMessage(
            f"Reference data ready — {len(df)} master rows. Open a file to validate.", 8000)

    def refresh_data(self):
        remote.force_refresh()
        self._load_reference_data(force=True)
        if self.user_df is not None:
            self.statusBar().showMessage("Reference data refreshed — re-validating…", 4000)

    # ---------- file open + validation ----------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx *.xls)")
        if not path:
            return
        try:
            self.user_df = pd.read_excel(path, dtype=str, header=0).reset_index(drop=True)
        except Exception as e:
            self._error("Could not read file", str(e)); return
        self._current_path = path
        self.setWindowTitle(f"Glasses Import Validator (Desktop) v{__version__} — {path.split('/')[-1]}")
        self.run_validation()

    def run_validation(self):
        if self.user_df is None:
            QMessageBox.information(self, "No file", "Open an Excel file first."); return
        if self.master_df is None:
            QMessageBox.information(self, "Please wait", "Reference data is still loading."); return
        self.statusBar().showMessage("Validating…")
        udf, mdf = self.user_df, self.master_df
        self._worker = Worker(lambda: vc.validate(udf, mdf))
        self._worker.done.connect(self._on_validated)
        self._worker.failed.connect(lambda e: self._error("Validation failed", e))
        self._worker.start()

    def _on_validated(self, result):
        self.result = result
        self.model = ValidationTableModel(self.user_df, result["cell_issues"])
        self.proxy.setSourceModel(self.model)
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
    def _error(self, title, detail):
        self.statusBar().showMessage(title, 6000)
        QMessageBox.critical(self, title, str(detail))


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
