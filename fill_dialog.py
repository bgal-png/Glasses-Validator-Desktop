"""Dialog for applying the ported Excel fill-macros to the sheet."""
from __future__ import annotations
from collections import Counter

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QCheckBox,
    QPushButton, QGroupBox, QLineEdit, QDialogButtonBox, QScrollArea, QWidget,
)

import rules


class FillDialog(QDialog):
    """Pick which fill rules to apply. Shows how many cells each would change."""

    def __init__(self, user_df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fill columns from rules")
        self.resize(620, 560)
        self.user_df = user_df
        self._boxes = {}

        root = QVBoxLayout(self)
        root.addWidget(QLabel(
            "These are the Excel fill macros. <b>Empty</b> = cells the rule can fill; "
            "<b>differs</b> = cells whose value disagrees with the rule."))

        # --- Name private (needs the three numbers the VBA form asked for) ---
        grp_pn = QGroupBox("Name private  —  a_PrivateNames (leave blank to skip a type)")
        g = QGridLayout(grp_pn)
        self.ed_eye = QLineEdit(); self.ed_sun = QLineEdit(); self.ed_comp = QLineEdit()
        for i, (lbl, ed, hint) in enumerate([
            ("Eyeglasses no.", self.ed_eye, "(Eyeglasses N)"),
            ("Sunglasses no.", self.ed_sun, "(Sunglasses N)"),
            ("Computer no.", self.ed_comp, "(Eyeglasses PC N)"),
        ]):
            g.addWidget(QLabel(lbl), i, 0); g.addWidget(ed, i, 1)
            g.addWidget(QLabel(hint), i, 2)
        self.chk_pn = QCheckBox("Apply Name private rule")
        g.addWidget(self.chk_pn, 3, 0, 1, 3)
        root.addWidget(grp_pn)
        for ed in (self.ed_eye, self.ed_sun, self.ed_comp):
            ed.textChanged.connect(self._refresh_counts)

        # --- Per-rule checkboxes ---
        grp = QGroupBox("Rules")
        gl = QGridLayout(grp)
        gl.addWidget(QLabel("<b>Rule</b>"), 0, 0)
        gl.addWidget(QLabel("<b>Empty</b>"), 0, 1)
        gl.addWidget(QLabel("<b>Differs</b>"), 0, 2)
        gl.addWidget(QLabel("<b>Source</b>"), 0, 3)
        self._count_labels = {}
        for r, (rid, label, _tk, _fn, src) in enumerate(rules.RULES, start=1):
            cb = QCheckBox(label); cb.setChecked(True)
            self._boxes[rid] = cb
            lf = QLabel("—"); ld = QLabel("—")
            self._count_labels[rid] = (lf, ld)
            gl.addWidget(cb, r, 0); gl.addWidget(lf, r, 1); gl.addWidget(ld, r, 2)
            gl.addWidget(QLabel(f"<i>{src}</i>"), r, 3)
        sa = QScrollArea(); sa.setWidgetResizable(True); sa.setWidget(grp)
        root.addWidget(sa, 1)

        # --- Options ---
        self.chk_overwrite = QCheckBox(
            "Also overwrite values that differ from the rule "
            "(the Excel macros always overwrite)")
        root.addWidget(self.chk_overwrite)
        self.chk_overwrite.toggled.connect(self._refresh_summary)

        row = QHBoxLayout()
        btn_all = QPushButton("Select all"); btn_none = QPushButton("Select none")
        btn_all.clicked.connect(lambda: self._set_all(True))
        btn_none.clicked.connect(lambda: self._set_all(False))
        row.addWidget(btn_all); row.addWidget(btn_none); row.addStretch(1)
        root.addLayout(row)

        self.lbl_summary = QLabel("—")
        self.lbl_summary.setStyleSheet("font-weight:bold; padding:4px;")
        root.addWidget(self.lbl_summary)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Apply")
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        root.addWidget(bb)

        for cb in self._boxes.values():
            cb.toggled.connect(self._refresh_summary)
        self.chk_pn.toggled.connect(self._refresh_summary)
        self._refresh_counts()

    # ---- helpers ----
    def _set_all(self, on):
        for cb in self._boxes.values():
            cb.setChecked(on)

    def private_params(self):
        if not self.chk_pn.isChecked():
            return None
        return {"eye": self.ed_eye.text().strip(),
                "sun": self.ed_sun.text().strip(),
                "comp": self.ed_comp.text().strip()}

    def _refresh_counts(self):
        results = rules.evaluate(self.user_df, private_params=self.private_params())
        self._results = results
        per = {}
        for r in results:
            c = per.setdefault(r["rule"], Counter())
            c[r["status"]] += 1
        for rid, (lf, ld) in self._count_labels.items():
            c = per.get(rid, Counter())
            lf.setText(str(c["fill"])); ld.setText(str(c["differs"]))
        self._pn_counts = per.get("private_name", Counter())
        self._refresh_summary()

    def _refresh_summary(self):
        sel = self.selected_rule_ids()
        n_fill = n_diff = 0
        for r in getattr(self, "_results", []):
            rid = r["rule"]
            if rid == "private_name":
                if not self.chk_pn.isChecked():
                    continue
            elif rid not in sel:
                continue
            if r["status"] == "fill":
                n_fill += 1
            elif r["status"] == "differs":
                n_diff += 1
        total = n_fill + (n_diff if self.chk_overwrite.isChecked() else 0)
        extra = f" (+{n_diff} differing left unchanged)" if (n_diff and not self.chk_overwrite.isChecked()) else ""
        self.lbl_summary.setText(f"Will change {total} cell(s){extra}")

    def selected_rule_ids(self):
        return {rid for rid, cb in self._boxes.items() if cb.isChecked()}

    def overwrite(self):
        return self.chk_overwrite.isChecked()
