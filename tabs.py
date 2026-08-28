"""The non-validation views: Image checker, Banned brands, Image renamer,
Syntax & duplicates. Each is a self-contained QWidget given the current
DataFrame via set_dataframe().
"""
from __future__ import annotations
import os
import shutil
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QCheckBox, QFileDialog, QMessageBox, QGroupBox, QTableWidget,
    QTableWidgetItem, QPlainTextEdit, QTabWidget, QSpinBox, QComboBox, QSlider,
    QProgressBar, QHeaderView, QAbstractItemView,
)

import checks
import renamer


# ---------------------------------------------------------------- helpers
class Job(QThread):
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


def _table(headers, rows, stretch_last=True):
    t = QTableWidget(len(rows), len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setAlternatingRowColors(True)
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            t.setItem(r, c, QTableWidgetItem("" if v is None else str(v)))
    t.resizeColumnsToContents()
    if stretch_last:
        t.horizontalHeader().setStretchLastSection(True)
    return t


def _metric_row(pairs):
    w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0, 0, 0, 0)
    labels = {}
    for key, name in pairs:
        box = QGroupBox(name); v = QVBoxLayout(box)
        lbl = QLabel("—"); lbl.setStyleSheet("font-size:18px; font-weight:bold;")
        lbl.setAlignment(Qt.AlignCenter)
        v.addWidget(lbl); h.addWidget(box)
        labels[key] = lbl
    return w, labels


class BaseTab(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self._job = None

    def set_dataframe(self, df):
        self.df = df
        self.on_dataframe(df)

    def on_dataframe(self, df):
        pass

    def _need_file(self):
        if self.df is None:
            QMessageBox.information(self, "No file", "Open an Excel file first.")
            return True
        return False


# ---------------------------------------------------------------- image checker
class ImageCheckerTab(BaseTab):
    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "<b>Image checker</b> — compares the image files you point at with the "
            "<i>Glasses name</i> column: which names have no image, which images "
            "match nothing, and any duplicates."))

        row = QHBoxLayout()
        self.btn_folder = QPushButton("📁 Select image folder…")
        self.btn_folder.clicked.connect(self.pick_folder)
        self.chk_recursive = QCheckBox("include subfolders"); self.chk_recursive.setChecked(True)
        self.lbl_folder = QLabel("<i>no folder selected</i>")
        row.addWidget(self.btn_folder); row.addWidget(self.chk_recursive)
        row.addWidget(self.lbl_folder, 1)
        lay.addLayout(row)

        self.paste = QPlainTextEdit()
        self.paste.setPlaceholderText(
            "…or paste file paths / names here, one per line "
            "(Ctrl+A in Explorer → right-click → Copy as path)")
        self.paste.setMaximumHeight(90)
        lay.addWidget(self.paste)

        b = QHBoxLayout()
        self.btn_check = QPushButton("🔍 Check images"); self.btn_check.clicked.connect(self.run_check)
        b.addWidget(self.btn_check); b.addStretch(1)
        lay.addLayout(b)

        mrow, self.metrics = _metric_row([
            ("excel", "Names in file"), ("images", "Images found"),
            ("matched", "Matched"), ("missing", "Missing image"),
        ])
        lay.addWidget(mrow)

        self.results = QTabWidget()
        lay.addWidget(self.results, 1)
        self._folder = None
        self._files = []

    def pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Select folder with images")
        if not d:
            return
        self._folder = d
        self._files = checks.collect_images(d, self.chk_recursive.isChecked())
        self.lbl_folder.setText(f"<b>{d}</b> — {len(self._files)} image(s)")

    def _image_names(self):
        names = list(self._files)
        pasted = [l.strip() for l in self.paste.toPlainText().split("\n") if l.strip()]
        return names + pasted

    def run_check(self):
        if self._need_file():
            return
        names = self._image_names()
        if not names:
            QMessageBox.information(self, "No images",
                                    "Select a folder or paste some file paths first.")
            return
        name_col = next((c for c in self.df.columns if "glasses name" in c.lower()),
                        self.df.columns[0])
        r = checks.check_images(self.df, name_col, names)
        self.metrics["excel"].setText(f"{r['excel_total']}")
        self.metrics["images"].setText(f"{r['images_total']}")
        self.metrics["matched"].setText(f"{r['matched']} / {r['excel_unique']}")
        self.metrics["missing"].setText(f"{len(r['missing'])}")

        self.results.clear()
        self.results.addTab(_table(["Excel row", "Name (no image found)"],
                                   [(i + 2, o) for i, o in r["missing"]]),
                            f"❌ Missing ({len(r['missing'])})")
        self.results.addTab(_table(["Image with no matching name"],
                                   [(x,) for x in r["extra"]]),
                            f"⚠️ Extra ({len(r['extra'])})")
        self.results.addTab(_table(["Image", "Times seen"],
                                   sorted(r["image_duplicates"].items())),
                            f"🔁 Duplicate images ({len(r['image_duplicates'])})")
        self.results.addTab(_table(["Name", "Occurrences"],
                                   sorted(r["excel_duplicates"].items())),
                            f"🔁 Duplicate names ({len(r['excel_duplicates'])})")


# ---------------------------------------------------------------- banned brands
class BannedBrandsTab(BaseTab):
    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "<b>Banned brands</b> — 🚫 marks a brand that must not be uploaded to "
            "that site. Alensa.ua is an allow-list, so anything outside its "
            "permitted brands is flagged."))
        mrow, self.metrics = _metric_row([
            ("brands", "Brands flagged"), ("sites", "Sites affected"),
            ("blocks", "Brand×site blocks"),
        ])
        lay.addWidget(mrow)
        self.lbl_status = QLabel("—"); self.lbl_status.setWordWrap(True)
        lay.addWidget(self.lbl_status)
        self.table = QTableWidget(0, 0)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        lay.addWidget(self.table, 1)
        self.lbl_clean = QLabel(""); self.lbl_clean.setWordWrap(True)
        lay.addWidget(self.lbl_clean)

    def on_dataframe(self, df):
        if df is None:
            return
        col = checks.find_brand_column(list(df.columns))
        if not col:
            self.lbl_status.setText("❌ No Brand column found in the file.")
            return
        r = checks.check_banned_brands(df, col)
        self.metrics["brands"].setText(str(len(r["flagged_brands"])))
        self.metrics["sites"].setText(f"{len(r['sites_with_flags'])} / "
                                      f"{len(r['flagged_by_site'])}")
        self.metrics["blocks"].setText(str(r["total_blocks"]))
        self.lbl_status.setText(
            f"Using column <code>{col}</code> — {len(r['brands'])} brand(s) in file: "
            + ", ".join(r["brands"][:20]) + (" …" if len(r["brands"]) > 20 else ""))

        brands, sites = r["flagged_brands"], r["sites_with_flags"]
        self.table.clear()
        self.table.setRowCount(len(sites)); self.table.setColumnCount(len(brands))
        self.table.setHorizontalHeaderLabels(brands)
        self.table.setVerticalHeaderLabels(sites)
        for i, site in enumerate(sites):
            for j, brand in enumerate(brands):
                blocked = brand in r["flagged_by_site"][site]
                it = QTableWidgetItem("🚫" if blocked else "")
                it.setTextAlignment(Qt.AlignCenter)
                if blocked:
                    it.setBackground(Qt.red)
                    it.setForeground(Qt.white)
                self.table.setItem(i, j, it)
        self.table.resizeColumnsToContents()
        self.lbl_clean.setText("✅ No issues on: " + ", ".join(r["clean_sites"])
                               if r["clean_sites"] else "")


# ---------------------------------------------------------------- image renamer
class ImageRenamerTab(BaseTab):
    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "<b>Image renamer</b> — renames product images to the canonical names "
            "from the <i>Glasses name</i> column (barcode-named files are matched "
            "via the Barcode column), optionally resizing them to a centred PNG."))

        row = QHBoxLayout()
        self.btn_src = QPushButton("📁 Source folder…"); self.btn_src.clicked.connect(self.pick_src)
        self.lbl_src = QLabel("<i>none</i>")
        row.addWidget(self.btn_src); row.addWidget(self.lbl_src, 1)
        lay.addLayout(row)

        row2 = QHBoxLayout()
        self.btn_out = QPushButton("📂 Output folder…"); self.btn_out.clicked.connect(self.pick_out)
        self.lbl_out = QLabel("<i>none</i>")
        row2.addWidget(self.btn_out); row2.addWidget(self.lbl_out, 1)
        lay.addLayout(row2)

        opt = QGroupBox("Resize (optional)")
        g = QGridLayout(opt)
        self.chk_resize = QCheckBox("Resize to centred PNG")
        self.sp_w = QSpinBox(); self.sp_w.setRange(100, 10000); self.sp_w.setValue(2400); self.sp_w.setSuffix(" px")
        self.sp_h = QSpinBox(); self.sp_h.setRange(100, 10000); self.sp_h.setValue(1800); self.sp_h.setSuffix(" px")
        self.cb_bg = QComboBox(); self.cb_bg.addItems(["auto", "white", "transparent"])
        self.sl_margin = QSlider(Qt.Horizontal); self.sl_margin.setRange(0, 25); self.sl_margin.setValue(5)
        self.lbl_margin = QLabel("5 %")
        self.sl_margin.valueChanged.connect(lambda v: self.lbl_margin.setText(f"{v} %"))
        g.addWidget(self.chk_resize, 0, 0, 1, 4)
        g.addWidget(QLabel("Width"), 1, 0); g.addWidget(self.sp_w, 1, 1)
        g.addWidget(QLabel("Height"), 1, 2); g.addWidget(self.sp_h, 1, 3)
        g.addWidget(QLabel("Background"), 2, 0); g.addWidget(self.cb_bg, 2, 1)
        g.addWidget(QLabel("Margin"), 2, 2)
        mrow = QHBoxLayout(); mrow.addWidget(self.sl_margin); mrow.addWidget(self.lbl_margin)
        mw = QWidget(); mw.setLayout(mrow); g.addWidget(mw, 2, 3)
        lay.addWidget(opt)

        b = QHBoxLayout()
        self.btn_preview = QPushButton("🔍 Preview matches"); self.btn_preview.clicked.connect(self.preview)
        self.btn_apply = QPushButton("✅ Rename into output folder"); self.btn_apply.clicked.connect(self.apply)
        self.btn_apply.setEnabled(False)
        b.addWidget(self.btn_preview); b.addWidget(self.btn_apply); b.addStretch(1)
        lay.addLayout(b)

        mr, self.metrics = _metric_row([
            ("total", "Images"), ("matched", "Matched"),
            ("unmatched", "Unmatched"), ("collisions", "Collisions"),
        ])
        lay.addWidget(mr)
        self.progress = QProgressBar(); self.progress.setVisible(False)
        lay.addWidget(self.progress)
        self.results = QTabWidget(); lay.addWidget(self.results, 1)

        self._src = self._out = None
        self._plan = []

    def pick_src(self):
        d = QFileDialog.getExistingDirectory(self, "Folder with the images to rename")
        if d:
            self._src = d
            n = len(checks.collect_images(d, False))
            self.lbl_src.setText(f"<b>{d}</b> — {n} image(s)")

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "Where to write the renamed images")
        if d:
            self._out = d
            self.lbl_out.setText(f"<b>{d}</b>")

    def _entries_and_barcodes(self):
        name_col = next((c for c in self.df.columns if "glasses name" in c.lower()),
                        self.df.columns[0])
        entries = []
        for v in self.df[name_col].dropna().astype(str):
            e = renamer.parse_list_entry(v.strip())
            if e:
                entries.append(e)
        bcol = next((c for c in self.df.columns if "barcode" in c.lower()), None)
        bmap = {}
        if bcol:
            import re as _re
            for _i, row in self.df.iterrows():
                bc = _re.sub(r"\D", "", str(row[bcol]))
                nm = str(row[name_col]).strip()
                if bc and nm.lower() not in ("nan", "", "none"):
                    bmap.setdefault(bc, nm)
        return entries, bmap

    def preview(self):
        if self._need_file():
            return
        if not self._src:
            QMessageBox.information(self, "No source", "Pick the source folder first."); return
        files = checks.collect_images(self._src, False)
        if not files:
            QMessageBox.information(self, "No images", "That folder has no images."); return
        entries, bmap = self._entries_and_barcodes()
        if not entries:
            QMessageBox.information(self, "No names", "No usable names in the file."); return

        plan = []
        for f in files:
            fname = os.path.basename(f)
            res = renamer.match_filename(fname, entries, bmap)
            row = {"source": fname, "path": f, "status": res["status"]}
            if res["status"] == "matched":
                ext = ".png" if self.chk_resize.isChecked() else Path(fname).suffix
                row["target"] = renamer.target_name_for(res["entry"], ext)
                row["leftover_tokens"] = res.get("leftover_tokens", [])
            else:
                row["target"] = None
                row["reason"] = res
            plan.append(row)
        renamer.resolve_collisions(plan)
        self._plan = plan

        matched = [r for r in plan if r["status"] == "matched"]
        unmatched = [r for r in plan if r["status"] != "matched"]
        self.metrics["total"].setText(str(len(plan)))
        self.metrics["matched"].setText(str(len(matched)))
        self.metrics["unmatched"].setText(str(len(unmatched)))
        self.metrics["collisions"].setText(str(sum(1 for r in plan if r.get("collision"))))

        self.results.clear()
        self.results.addTab(_table(["Source", "→", "Target", "Note"],
                                   [(r["source"], "→", r["target"], r.get("collision", ""))
                                    for r in matched]),
                            f"✅ Matched ({len(matched)})")
        self.results.addTab(_table(["Source", "Reason"],
                                   [(r["source"], r.get("reason", {}).get("status", "?"))
                                    for r in unmatched]),
                            f"⚠️ Unmatched ({len(unmatched)})")
        self.btn_apply.setEnabled(bool(matched))

    def apply(self):
        matched = [r for r in self._plan if r["status"] == "matched"]
        if not matched:
            QMessageBox.information(self, "Nothing to do", "Run a preview first."); return
        if not self._out:
            QMessageBox.information(self, "No output folder",
                                    "Pick the output folder first."); return
        if QMessageBox.question(
                self, "Rename images",
                f"Write {len(matched)} renamed image(s) into:\n{self._out}\n\n"
                "Source files are left untouched.") != QMessageBox.Yes:
            return

        do_resize = self.chk_resize.isChecked()
        w, h = self.sp_w.value(), self.sp_h.value()
        bg = self.cb_bg.currentText()
        margin = self.sl_margin.value() / 100.0
        errors = []
        self.progress.setVisible(True); self.progress.setRange(0, len(matched))
        for i, r in enumerate(matched, 1):
            self.progress.setValue(i)
            dest = os.path.join(self._out, r["target"])
            try:
                if do_resize:
                    with open(r["path"], "rb") as fh:
                        data = renamer.resize_centered(fh.read(), w, h, margin, bg)
                    with open(dest, "wb") as fh:
                        fh.write(data)
                else:
                    shutil.copy2(r["path"], dest)
            except Exception as e:
                errors.append((r["source"], str(e)))
        self.progress.setVisible(False)
        msg = f"Wrote {len(matched) - len(errors)} image(s) to:\n{self._out}"
        if errors:
            msg += f"\n\n{len(errors)} failed:\n" + "\n".join(f"  {s}: {e}" for s, e in errors[:8])
        QMessageBox.information(self, "Done", msg)


# ---------------------------------------------------------------- syntax & dups
class SyntaxDuplicatesTab(BaseTab):
    def __init__(self, get_names):
        super().__init__()
        self._get_names = get_names
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(
            "<b>Syntax &amp; duplicates</b> — checks the <i>Glasses name</i> column "
            "against the name master: names that already exist, names repeated "
            "inside this file, and naming patterns never seen before."))
        b = QHBoxLayout()
        self.btn_run = QPushButton("🧬 Analyse"); self.btn_run.clicked.connect(self.run)
        self.btn_export = QPushButton("⬇ Export duplicates"); self.btn_export.clicked.connect(self.export)
        self.btn_export.setEnabled(False)
        b.addWidget(self.btn_run); b.addWidget(self.btn_export); b.addStretch(1)
        lay.addLayout(b)
        mr, self.metrics = _metric_row([
            ("master", "In master"), ("infile", "In-file dups"), ("syntax", "Odd syntax"),
        ])
        lay.addWidget(mr)
        self.lbl_info = QLabel("—"); lay.addWidget(self.lbl_info)
        self.table = QTableWidget(0, 4); lay.addWidget(self.table, 1)
        self.table.setHorizontalHeaderLabels(["Excel row", "Name", "Issue", "Detail"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._result = None

    def run(self):
        if self._need_file():
            return
        names = self._get_names()
        if not names:
            QMessageBox.information(self, "Name master unavailable",
                                    "The name master hasn't loaded yet."); return
        col = next((c for c in self.df.columns if "glasses name" in c.lower()),
                   self.df.columns[0])
        r = checks.check_syntax_duplicates(self.df, col, names)
        self._result = r
        rep = r["report"]
        counts = {k: sum(1 for x in rep if x["issue"] == k)
                  for k in ("DUPLICATE", "IN-FILE DUPLICATE", "SUSPICIOUS SYNTAX")}
        self.metrics["master"].setText(str(counts["DUPLICATE"]))
        self.metrics["infile"].setText(str(counts["IN-FILE DUPLICATE"]))
        self.metrics["syntax"].setText(str(counts["SUSPICIOUS SYNTAX"]))
        self.lbl_info.setText(f"Compared against {len(names):,} master names — "
                              f"{len(rep)} issue(s) found.")
        self.table.setRowCount(len(rep))
        for i, x in enumerate(rep):
            for c, v in enumerate([x["row"] + 2, x["name"], x["issue"], x["detail"]]):
                it = QTableWidgetItem(str(v))
                if x["issue"] == "DUPLICATE":
                    it.setBackground(Qt.red); it.setForeground(Qt.white)
                elif x["issue"] == "IN-FILE DUPLICATE":
                    it.setBackground(Qt.darkYellow)
                self.table.setItem(i, c, it)
        self.table.resizeColumnsToContents()
        self.btn_export.setEnabled(bool(r["duplicate_rows"]))

    def export(self):
        if not self._result or not self._result["duplicate_rows"]:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save duplicate rows",
                                              "duplicate_items.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.df.loc[self._result["duplicate_rows"]].to_excel(path, index=False)
            QMessageBox.information(self, "Exported", f"Saved → {path}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))
