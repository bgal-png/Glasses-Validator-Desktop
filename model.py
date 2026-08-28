"""Qt table model that renders the user's spreadsheet and colours flagged cells."""
from __future__ import annotations
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel
from PySide6.QtGui import QColor

ERROR_BG = QColor("#ffb3b3")     # red-ish
WARNING_BG = QColor("#ffe08a")   # amber
ERROR_TYPES = {"empty_required", "empty_required_sun", "meta_format", "invalid_content"}


class ValidationTableModel(QAbstractTableModel):
    def __init__(self, df, cell_issues=None):
        super().__init__()
        self._df = df.reset_index(drop=True)
        self._cell_issues = cell_issues or {}
        self._cols = list(self._df.columns)
        self._rows_with_issues = {r for (r, _c) in self._cell_issues}

    # ---- data plumbing ----
    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._cols)

    def _issues_at(self, row, col_name):
        return self._cell_issues.get((row, col_name), [])

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        col_name = self._cols[col]
        if role in (Qt.DisplayRole, Qt.EditRole):
            v = self._df.iat[row, col]
            return "" if v is None or (isinstance(v, float)) and str(v) == "nan" else str(v)
        issues = self._issues_at(row, col_name)
        if not issues:
            return None
        if role == Qt.BackgroundRole:
            has_error = any(i["type"] in ERROR_TYPES for i in issues)
            return ERROR_BG if has_error else WARNING_BG
        if role == Qt.ToolTipRole:
            return "\n".join(f"• {i['message']}" for i in issues)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self._cols[section]
        return str(section + 2)  # Excel row number (header row is 1)

    # ---- helpers used by the UI ----
    def row_has_issue(self, row):
        return row in self._rows_with_issues

    def cell_issue_types(self, row, col_name):
        return {i["type"] for i in self._issues_at(row, col_name)}

    def dataframe(self):
        return self._df

    def issue_rows_sorted(self):
        return sorted(self._rows_with_issues)


class IssueFilterProxy(QSortFilterProxyModel):
    """Optionally show only rows that contain at least one issue."""
    def __init__(self):
        super().__init__()
        self._only_issues = False

    def set_only_issues(self, on: bool):
        self._only_issues = on
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        if not self._only_issues:
            return True
        model = self.sourceModel()
        return model.row_has_issue(source_row)
