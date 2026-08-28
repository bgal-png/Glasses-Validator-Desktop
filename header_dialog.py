"""Review dialog for proposed column-header repairs."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QDialogButtonBox, QAbstractItemView, QHeaderView,
)

METHOD_HELP = {
    "normalised": "same text, different spacing/punctuation",
    "no-id": "same name, the 'ID: nn' part was missing",
    "similar": "very close spelling",
    "position": "header carried no information — matched by column position",
}


class HeaderFixDialog(QDialog):
    def __init__(self, proposals, unmatched_file, unmatched_tpl, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fix column headers")
        self.resize(880, 520)
        self.proposals = proposals

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "These headers don't match the import template. Renaming them lets the "
            "validator map every column.<br>"
            "<b>Untick anything you don't want changed</b> — rows matched only by "
            "column position are highlighted."))

        self.table = QTableWidget(len(proposals), 5)
        self.table.setHorizontalHeaderLabels(
            ["Apply", "Current header", "→", "Correct header", "Why"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        for r, p in enumerate(proposals):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked)
            self.table.setItem(r, 0, chk)
            why = METHOD_HELP.get(p["method"], p["method"])
            if p["method"] != "position":
                why = f"{why}"
            for c, v in enumerate([p["current"] or "(blank)", "→", p["proposed"], why], start=1):
                it = QTableWidgetItem(str(v))
                if p["method"] == "position":
                    it.setBackground(Qt.darkYellow if c != 2 else Qt.transparent)
                self.table.setItem(r, c, it)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.table, 1)

        notes = []
        if unmatched_tpl:
            notes.append("⚠️ Template columns with no match in the file: "
                         + ", ".join(unmatched_tpl))
        if unmatched_file:
            notes.append("ℹ️ File columns not in the template (left alone): "
                         + ", ".join(str(u) for u in unmatched_file))
        if notes:
            lbl = QLabel("<br>".join(notes)); lbl.setWordWrap(True)
            lay.addWidget(lbl)

        row = QHBoxLayout()
        ba = QPushButton("Select all"); bn = QPushButton("Select none")
        ba.clicked.connect(lambda: self._set_all(True))
        bn.clicked.connect(lambda: self._set_all(False))
        row.addWidget(ba); row.addWidget(bn); row.addStretch(1)
        lay.addLayout(row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Apply renames")
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _set_all(self, on):
        for r in range(self.table.rowCount()):
            self.table.item(r, 0).setCheckState(Qt.Checked if on else Qt.Unchecked)

    def accepted_proposals(self):
        return [p for r, p in enumerate(self.proposals)
                if self.table.item(r, 0).checkState() == Qt.Checked]
