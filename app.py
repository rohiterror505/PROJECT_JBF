#!/usr/bin/env python3
"""
JAI BHADRA FOUNDATION - LUCKY DRAW COUPON MANAGER
PyQt6 desktop UI for rohit.py

RUN:
    python app.py

Reuses all coupon generation, Excel tracking, and validation logic from
rohit.py by importing it as a module. No duplicate business logic.
"""

import sys
import tempfile
from datetime import datetime

from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QSize, QTimer
)
from PyQt6.QtGui import (
    QPixmap, QImage, QColor, QFont, QFontDatabase, QAction
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout,
    QHBoxLayout, QFormLayout, QLabel, QLineEdit, QSpinBox, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QProgressBar,
    QStackedWidget, QScrollArea, QFrame, QMessageBox, QDialog, QCheckBox,
    QGroupBox, QSizePolicy, QSplitter
)

import rohit


# ============================================================
# THEME
# ============================================================

MAROON = "#C44404"
MAROON_LIGHT = "#E85D04"
GOLD = "#C99B2E"
GOLD_LIGHT = "#F3D264"
GOLD_PALE = "#FFF1BD"
CREAM = "#FFF9E8"
CREAM_2 = "#FFFDF6"
WHITE = "#FFFFFF"
GREY = "#666666"
BLACK = "#151515"

QSS = f"""
QMainWindow, QWidget {{
    background-color: {CREAM};
    color: {BLACK};
    font-family: "Georgia", "Cambria", serif;
    font-size: 14px;
}}
QTabWidget::pane {{
    border: 2px solid {GOLD};
    background: {CREAM_2};
    border-radius: 8px;
    top: -1px;
}}
QTabBar::tab {{
    background: {CREAM};
    color: {MAROON};
    border: 2px solid {GOLD};
    border-bottom: none;
    padding: 10px 22px;
    margin-right: 4px;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    font-weight: bold;
    min-width: 120px;
}}
QTabBar::tab:selected {{
    background: {MAROON};
    color: {GOLD_LIGHT};
}}
QTabBar::tab:hover:!selected {{
    background: {GOLD_PALE};
}}
QPushButton {{
    background: {MAROON};
    color: {GOLD_LIGHT};
    border: 2px solid {GOLD};
    padding: 9px 20px;
    border-radius: 8px;
    font-weight: bold;
    min-width: 100px;
}}
QPushButton:hover {{
    background: {MAROON_LIGHT};
    border-color: {GOLD_LIGHT};
}}
QPushButton:pressed {{
    background: {MAROON};
}}
QPushButton:disabled {{
    background: #BBBBBB;
    color: #888888;
    border-color: #CCCCCC;
}}
QPushButton[flat="true"] {{
    background: {CREAM_2};
    color: {MAROON};
    border: 1px solid {GOLD};
}}
QPushButton[flat="true"]:hover {{
    background: {GOLD_PALE};
}}
QLineEdit, QSpinBox, QComboBox {{
    background: {WHITE};
    border: 2px solid {GOLD};
    border-radius: 6px;
    padding: 6px 10px;
    selection-background-color: {GOLD_LIGHT};
    min-height: 22px;
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
    border: 2px solid {MAROON};
}}
QTableWidget {{
    background: {WHITE};
    alternate-background-color: {GOLD_PALE};
    border: 2px solid {GOLD};
    border-radius: 6px;
    gridline-color: #D8BD79;
    selection-background-color: {GOLD_LIGHT};
    selection-color: {BLACK};
}}
QHeaderView::section {{
    background: {rohit.MAROON_DARK if hasattr(rohit,'MAROON_DARK') else MAROON};
    color: {GOLD_LIGHT};
    font-weight: bold;
    padding: 8px;
    border: 1px solid {GOLD};
}}
QProgressBar {{
    border: 2px solid {GOLD};
    border-radius: 6px;
    background: {CREAM_2};
    text-align: center;
    height: 24px;
}}
QProgressBar::chunk {{
    background: {MAROON};
    border-radius: 4px;
}}
QGroupBox {{
    border: 2px solid {GOLD};
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px 10px 10px 10px;
    font-weight: bold;
    color: {MAROON};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 6px;
}}
QLabel[heading="true"] {{
    color: {MAROON};
    font-weight: bold;
    font-size: 16px;
}}
QLabel[stat="true"] {{
    color: {GOLD_LIGHT};
    font-weight: bold;
}}
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollBar:vertical {{
    background: {CREAM};
    width: 14px;
    border: 1px solid {GOLD};
    border-radius: 6px;
}}
QScrollBar::handle:vertical {{
    background: {GOLD};
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {GOLD_LIGHT};
}}
QMessageBox {{
    background: {CREAM_2};
}}
QMessageBox QLabel {{
    color: {BLACK};
    font-size: 14px;
}}
"""


def card(title, value, color=MAROON, bg=WHITE):
    """Return a styled stat card widget."""
    f = QFrame()
    f.setStyleSheet(f"""
        QFrame {{
            background: {bg};
            border: 2px solid {GOLD};
            border-radius: 10px;
        }}
    """)
    lay = QVBoxLayout(f)
    lay.setContentsMargins(18, 12, 18, 12)
    title_lbl = QLabel(title)
    title_lbl.setStyleSheet(f"color:{GREY}; font-size:12px; font-weight:bold;")
    val_lbl = QLabel(value)
    val_lbl.setStyleSheet(f"color:{color}; font-size:24px; font-weight:bold;")
    val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
    lay.addWidget(title_lbl)
    lay.addWidget(val_lbl)
    f.setMinimumHeight(96)
    return f


class NumericTableWidgetItem(QTableWidgetItem):
    """QTableWidgetItem that sorts by its numeric value instead of by the
    display string.  Used for the S.No / Start / End / Qty / Donation
    columns in the Sales table so that '10' sorts after '9' rather than
    after '1' (which is what lexicographic string compare would do)."""

    def __init__(self, value):
        super().__init__()
        try:
            self._num = float(value)
        except (TypeError, ValueError):
            self._num = 0.0

    def __lt__(self, other):
        if isinstance(other, NumericTableWidgetItem):
            return self._num < other._num
        try:
            return self._num < float(other.text())
        except (TypeError, ValueError):
            return super().__lt__(other)


# ============================================================
# WORKERS
# ============================================================

class SaleWorker(QThread):
    """Generate coupons + record a single sale in the background."""
    done = pyqtSignal(str, bool)  # message, success

    def __init__(self, name, phone, address, start_no, qty):
        super().__init__()
        self.name = name
        self.phone = phone
        self.address = address
        self.start_no = start_no
        self.qty = qty

    def run(self):
        try:
            end = self.start_no + self.qty - 1
            if end > rohit.MAX_COUPON:
                self.done.emit(
                    f"End number {end:04d} would exceed max {rohit.MAX_COUPON}.",
                    False,
                )
                return
            overlap = rohit.is_already_sold(self.start_no, end)
            if overlap:
                self.done.emit(
                    f"Coupons {overlap[0]:04d}-{overlap[1]:04d} are already sold.",
                    False,
                )
                return
            amount = self.qty * rohit.PRICE_PER_COUPON
            filename = rohit.create_coupon(
                self.start_no, end,
                buyer=self.name or None,
                phone=self.phone or None,
                address=self.address or None,
                amount=amount,
            )
            rohit.record_sale(
                self.name, self.phone, self.address, self.start_no, end, amount=amount
            )
            self.done.emit(
                f"Sale complete! Coupons {self.start_no:04d}-{end:04d} saved as {filename}.",
                True,
            )
        except RuntimeError as exc:
            self.done.emit(str(exc), False)
        except Exception as exc:
            self.done.emit(f"Error: {exc}", False)


class PreviewWorker(QThread):
    """Render a preview coupon in memory (no file is written to disk)."""
    done = pyqtSignal(object)  # QImage or None

    def __init__(self, start_no, end_no, name, phone, address, amount=None):
        super().__init__()
        self.start_no = start_no
        self.end_no = end_no
        self.name = name
        self.phone = phone
        self.address = address
        self.amount = amount

    def run(self):
        try:
            img = rohit._render_coupon(
                self.start_no, self.end_no,
                buyer=self.name or None,
                phone=self.phone or None,
                address=self.address or None,
                amount=self.amount,
            ).convert("RGB")
            data = img.tobytes()
            qimg = QImage(data, img.width, img.height, img.width * 3,
                          QImage.Format.Format_RGB888)
            # Keep the raw bytes alive for as long as the QImage exists,
            # otherwise the buffer is garbage-collected and the image data
            # becomes invalid / crashes when the GUI thread paints it.
            qimg._raw = data
            self.done.emit(qimg)
        except Exception:
            self.done.emit(None)


class PhysicalWorker(QThread):
    """Generate physical coupon sets in the background with progress."""
    progress = pyqtSignal(int, int, str)  # current, total, message
    done = pyqtSignal(str, bool)

    def __init__(self, set_size, start_no, num_sets):
        super().__init__()
        self.set_size = set_size
        self.start_no = start_no
        self.num_sets = num_sets

    def run(self):
        try:
            end_block = self.start_no + self.num_sets * self.set_size - 1
            if end_block > rohit.MAX_COUPON:
                self.done.emit(
                    f"Would end at {end_block:04d}, exceeds max {rohit.MAX_COUPON}.",
                    False,
                )
                return
            overlap = rohit.is_already_sold(self.start_no, end_block)
            if overlap:
                self.done.emit(
                    f"Coupons {overlap[0]:04d}-{overlap[1]:04d} already sold.",
                    False,
                )
                return
            generated = 0
            set_amount = rohit.price_for_set_size(self.set_size)
            for i in range(self.num_sets):
                s = self.start_no + i * self.set_size
                e = s + self.set_size - 1
                self.progress.emit(i, self.num_sets,
                                   f"Generating {s:04d}-{e:04d}...")
                try:
                    rohit.create_coupon(s, e, amount=set_amount)
                    rohit.record_sale(
                        None, None, None, s, e, sale_type="PHYSICAL",
                        set_size=self.set_size, amount=set_amount,
                    )
                    generated += 1
                except Exception as exc:
                    self.progress.emit(i, self.num_sets, f"Failed {s:04d}: {exc}")
            self.done.emit(
                f"Done. {generated}/{self.num_sets} set(s) generated. "
                f"Numbers {self.start_no:04d}-{end_block:04d} locked.",
                True,
            )
        except RuntimeError as exc:
            self.done.emit(str(exc), False)
        except Exception as exc:
            self.done.emit(f"Error: {exc}", False)


class DeleteWorker(QThread):
    """Run a delete operation in the background so the UI does not freeze
    while Excel is being saved and PNG files are being removed."""
    done = pyqtSignal(object, bool, str)  # caller_token, success, message

    def __init__(self, fn, *args, caller_token=None):
        super().__init__()
        self._fn = fn
        self._args = args
        self._caller_token = caller_token

    def run(self):
        try:
            result = self._fn(*self._args)
            if result is False:
                self.done.emit(self._caller_token, False, "No matching sale found.")
            elif isinstance(result, int):
                self.done.emit(self._caller_token, True, f"Deleted {result} row(s).")
            else:
                self.done.emit(self._caller_token, True, "Done.")
        except RuntimeError as exc:
            self.done.emit(self._caller_token, False, str(exc))
        except Exception as exc:
            self.done.emit(self._caller_token, False, f"Error: {exc}")


class UpdateWorker(QThread):
    """Run rohit.update_sale in the background so the UI stays responsive
    while Excel is being saved and the coupon PNG is re-rendered."""
    done = pyqtSignal(object, bool, str)  # caller_token, success, message

    def __init__(self, sno, name, phone, address, caller_token=None):
        super().__init__()
        self._sno = sno
        self._name = name
        self._phone = phone
        self._address = address
        self._caller_token = caller_token

    def run(self):
        try:
            ok = rohit.update_sale(
                self._sno, self._name, self._phone, self._address,
                regen_png=True,
            )
            if ok:
                self.done.emit(
                    self._caller_token, True,
                    f"Sale #{self._sno} updated.",
                )
            else:
                self.done.emit(
                    self._caller_token, False,
                    f"No sale #{self._sno} found.",
                )
        except RuntimeError as exc:
            self.done.emit(self._caller_token, False, str(exc))
        except Exception as exc:
            self.done.emit(self._caller_token, False, f"Error: {exc}")


class RefreshWorker(QThread):
    """Fetch sales rows + sold ranges in a background thread so the GUI
    thread never blocks on Excel / OneDrive file I/O during refreshes.

    The result (rows, ranges) is emitted once the load completes; the GUI
    thread then updates the tabs from the in-memory snapshot."""
    done = pyqtSignal(object, object, object)  # rows, ranges, error(str or None)

    def run(self):
        try:
            # Ensure the file exists / is not corrupt BEFORE reading it.
            # This runs off the GUI thread so a slow OneDrive lock does not
            # freeze the window.
            rohit.ensure_sales_file()
            rows = rohit.list_sales(print_it=False)
            ranges = rohit.get_sold_ranges()
            self.done.emit(rows, ranges, None)
        except RuntimeError as exc:
            self.done.emit(None, None, str(exc))
        except Exception as exc:
            self.done.emit(None, None, f"Error: {exc}")


class DrawActionWorker(QThread):
    """Run a single draw-related rohit function (draw_winners,
    clear_draw_results, or get_draw_results) off the GUI thread.

    Emits (action_tag, success, message, results_json) where results_json
    is the list of result dicts (only meaningful for draw / get_draw_results)."""
    done = pyqtSignal(str, bool, str, object)

    def __init__(self, action):
        super().__init__()
        self._action = action  # "draw" | "clear" | "load"

    def run(self):
        tag = self._action
        try:
            if tag == "draw":
                results = rohit.draw_winners()
                self.done.emit(tag, True, "Draw complete! 50 winners selected from all 9999 coupons (40 consolation + 10 main). Unsold winners marked UNSOLD.", results)
            elif tag == "clear":
                cleared = rohit.clear_draw_results()
                msg = ("Lucky Draw results cleared." if cleared
                       else "No results to clear.")
                self.done.emit(tag, True, msg, None)
            elif tag == "load":
                results = rohit.get_draw_results()
                self.done.emit(tag, True, "Loaded.", results)
            else:
                self.done.emit(tag, False, "Unknown action.", None)
        except RuntimeError as exc:
            self.done.emit(tag, False, str(exc), None)
        except Exception as exc:
            self.done.emit(tag, False, f"Error: {exc}", None)


# ============================================================
# DIALOGS
# ============================================================

class ImageViewerDialog(QDialog):
    def __init__(self, path, title="Coupon"):
        super().__init__()
        self.setWindowTitle(title)
        self.setStyleSheet(f"background:{CREAM};")
        lay = QVBoxLayout(self)
        lbl = QLabel()
        pix = QPixmap(str(path))
        if pix.isNull():
            lbl.setText("Could not load image.")
        else:
            scaled = pix.scaledToWidth(
                600, Qt.TransformationMode.SmoothTransformation
            )
            lbl.setPixmap(scaled)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        lay.addWidget(close, alignment=Qt.AlignmentFlag.AlignCenter)


class EditSaleDialog(QDialog):
    """Dialog to edit the buyer details (Name, Phone, Address) of an
    existing sale row.  Coupon numbers, qty, amount and type are not
    editable here — only the buyer fields, per the update_sale contract."""

    def __init__(self, sno, name="", phone="", address="", stype="SALE",
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Edit Sale #{sno}")
        self.setStyleSheet(f"background:{CREAM};")
        self._sno = sno

        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(20, 20, 20, 20)

        title = QLabel(f"Edit Buyer Details — Sale #{sno}"
                       + (f"  ({stype})" if stype != "SALE" else ""))
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        info = QLabel("Only the buyer details can be changed. Coupon numbers, "
                      "quantity and amount stay locked.")
        info.setWordWrap(True)
        info.setStyleSheet(f"color:{GREY}; font-size:12px;")
        lay.addWidget(info)

        form = QFormLayout()
        form.setSpacing(8)
        self.name_in = QLineEdit(name)
        self.name_in.setPlaceholderText("Buyer name")
        self.phone_in = QLineEdit(phone)
        self.phone_in.setPlaceholderText("Phone number")
        self.addr_in = QLineEdit(address)
        self.addr_in.setPlaceholderText("Address")
        form.addRow("Name:", self.name_in)
        form.addRow("Phone:", self.phone_in)
        form.addRow("Address:", self.addr_in)
        lay.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setProperty("flat", True)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        self.save_btn = QPushButton("Save Changes")
        self.save_btn.clicked.connect(self.accept)
        btn_row.addWidget(self.save_btn)
        lay.addLayout(btn_row)

    def values(self):
        return (
            self.name_in.text().strip(),
            self.phone_in.text().strip(),
            self.addr_in.text().strip(),
        )


# ============================================================
# TABS
# ============================================================

class DashboardTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self.build()
        # First refresh is deferred to MainWindow._startup_init so the window
        # paints before any Excel I/O happens.

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(14)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("DASHBOARD")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        self.grid = QHBoxLayout()
        self.grid.setSpacing(12)
        self.card_sold = card("TOTAL SOLD", "0")
        self.card_revenue = card("REVENUE", "₹0")
        self.card_remaining = card("REMAINING", str(rohit.MAX_COUPON), color=GOLD)
        self.card_gaps = card("GAPS", "0", color=MAROON_LIGHT)
        self.grid.addWidget(self.card_sold)
        self.grid.addWidget(self.card_revenue)
        self.grid.addWidget(self.card_remaining)
        self.grid.addWidget(self.card_gaps)
        lay.addLayout(self.grid)

        self.split = QHBoxLayout()
        self.split.setSpacing(12)
        self.card_sale = card("SALE ROWS", "0")
        self.card_phys = card("PHYSICAL ROWS", "0", color=MAROON_LIGHT)
        self.card_last = card("LAST COUPON", "-")
        self.split.addWidget(self.card_sale)
        self.split.addWidget(self.card_phys)
        self.split.addWidget(self.card_last)
        lay.addLayout(self.split)

        recent_title = QLabel("RECENT SALES")
        recent_title.setProperty("heading", True)
        lay.addWidget(recent_title)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "S.No", "Name", "Phone", "Start", "End", "Qty", "Date"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for i in [0, 2, 3, 4, 5, 6]:
            hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.table)

        sample_group = QGroupBox("Generate Sample Coupon (no database entry)")
        sample_lay = QHBoxLayout(sample_group)
        sample_lay.addWidget(QLabel("Coupon No:"))
        self.sample_no_in = QSpinBox()
        self.sample_no_in.setRange(0, 9999)
        self.sample_no_in.setValue(0)
        sample_lay.addWidget(self.sample_no_in)
        self.sample_btn = QPushButton("Generate Sample Coupon")
        self.sample_btn.clicked.connect(self.generate_sample)
        sample_lay.addWidget(self.sample_btn)
        sample_lay.addStretch()
        lay.addWidget(sample_group)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setProperty("flat", True)
        refresh_btn.clicked.connect(self.refresh)
        lay.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignRight)

        lay.addStretch()

    def generate_sample(self):
        num = self.sample_no_in.value()
        try:
            img = rohit._render_coupon(num, num, amount=rohit.PRICE_PER_COUPON, sample=True)
        except Exception as exc:
            QMessageBox.warning(self, "Error", f"Could not render sample: {exc}")
            return
        rohit.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        path = rohit.OUTPUT_DIR / f"sample_{num:04d}.png"
        try:
            img.save(str(path), "PNG", optimize=True, dpi=(300, 300))
        except Exception as exc:
            QMessageBox.warning(self, "Error", f"Could not save sample: {exc}")
            return
        dlg = ImageViewerDialog(path, f"Sample Coupon {num:04d}")
        dlg.exec()

    def refresh(self):
        # Ask the main window for a fresh snapshot; actual disk I/O happens
        # in a background thread there, so this tab never blocks.
        self.parent_window.request_refresh()

    def apply_snapshot(self, rows, ranges):
        sale_count = sum(1 for r in rows
                         if (r[8] if len(r) >= 9 else "SALE") == "SALE")
        phys_count = sum(1 for r in rows
                         if (r[8] if len(r) >= 9 else "SALE") == "PHYSICAL")
        total_qty = sum((r[6] if len(r) >= 7 else 0) or 0 for r in rows)
        # Revenue = sum of each row's stored donation amount (set-based
        # pricing).  Fall back to the legacy qty*100 for old rows.
        def _row_amount(r):
            if len(r) >= 11 and r[10] is not None:
                return r[10] or 0
            qty = (r[6] if len(r) >= 7 else 0) or 0
            stype = r[8] if len(r) >= 9 else "SALE"
            return (rohit.price_for_set_size(qty) if stype == "PHYSICAL"
                    else qty * rohit.PRICE_PER_COUPON)
        revenue = sum(_row_amount(r) for r in rows)

        assigned = set()
        for s, e in ranges:
            for n in range(s, e + 1):
                assigned.add(n)
        gaps = 0
        if assigned:
            low, high = min(assigned), max(assigned)
            gaps = len([n for n in range(low, high + 1) if n not in assigned])

        self.card_sold.findChildren(QLabel)[1].setText(str(total_qty))
        self.card_revenue.findChildren(QLabel)[1].setText(f"₹{revenue:,}")
        self.card_remaining.findChildren(QLabel)[1].setText(
            str(rohit.MAX_COUPON - total_qty)
        )
        self.card_gaps.findChildren(QLabel)[1].setText(str(gaps))
        self.card_sale.findChildren(QLabel)[1].setText(str(sale_count))
        self.card_phys.findChildren(QLabel)[1].setText(str(phys_count))

        if rows:
            last = rows[-1]
            start = last[4]
            end = last[5]
            label = f"{start:04d}" if start == end else f"{start:04d}-{end:04d}"
            self.card_last.findChildren(QLabel)[1].setText(label)
        else:
            self.card_last.findChildren(QLabel)[1].setText("-")

        recent = rows[-5:] if len(rows) > 5 else rows
        self.table.setRowCount(len(recent))
        for i, r in enumerate(recent):
            vals = [str(r[0] or ""), str(r[1] or ""), str(r[2] or ""),
                    str(r[4] or ""), str(r[5] or ""), str(r[6] or ""),
                    str(r[7] or "")]
            for c, v in enumerate(vals):
                self.table.setItem(i, c, QTableWidgetItem(v))
        self.table.resizeColumnsToContents()
        if self.table.columnWidth(1) < 120:
            self.table.setColumnWidth(1, 120)


class NewSaleTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.preview_worker = None
        self.preview_path = None
        self.build()

    def build(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(12)
        outer.setContentsMargins(18, 18, 18, 18)

        title = QLabel("NEW SALE")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(title)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        form_container = QWidget()
        form_lay = QVBoxLayout(form_container)
        form_lay.setSpacing(10)

        group = QGroupBox("Buyer Details")
        form = QFormLayout(group)
        form.setSpacing(8)
        self.name_in = QLineEdit()
        self.name_in.setPlaceholderText("Buyer name")
        self.phone_in = QLineEdit()
        self.phone_in.setPlaceholderText("Phone number")
        self.addr_in = QLineEdit()
        self.addr_in.setPlaceholderText("Address")
        form.addRow("Name:", self.name_in)
        form.addRow("Phone:", self.phone_in)
        form.addRow("Address:", self.addr_in)
        form_lay.addWidget(group)

        coup_group = QGroupBox("Coupon Numbers")
        coup_form = QFormLayout(coup_group)
        coup_form.setSpacing(8)
        self.start_in = QSpinBox()
        self.start_in.setRange(1, rohit.MAX_COUPON)
        self.start_in.setValue(1)
        self.qty_in = QSpinBox()
        self.qty_in.setRange(1, rohit.MAX_COUPON)
        self.qty_in.setValue(1)
        self.end_label = QLabel("0001")
        self.end_label.setStyleSheet(f"color:{MAROON}; font-weight:bold; font-size:18px;")
        coup_form.addRow("Start No:", self.start_in)
        coup_form.addRow("Quantity:", self.qty_in)
        coup_form.addRow("End No:", self.end_label)
        form_lay.addWidget(coup_group)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet(f"color:{MAROON}; font-weight:bold;")
        form_lay.addWidget(self.status_label)

        self.sell_btn = QPushButton("Generate & Record Sale")
        self.sell_btn.setMinimumHeight(46)
        self.sell_btn.clicked.connect(self.do_sale)
        form_lay.addWidget(self.sell_btn)

        form_lay.addStretch()
        splitter.addWidget(form_container)

        preview_container = QWidget()
        prev_lay = QVBoxLayout(preview_container)
        prev_lay.setContentsMargins(6, 0, 6, 0)
        prev_title = QLabel("LIVE PREVIEW")
        prev_title.setProperty("heading", True)
        prev_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        prev_lay.addWidget(prev_title)
        self.preview_label = QLabel("Enter details to see a preview")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet(
            f"background:{WHITE}; border:2px solid {GOLD}; border-radius:8px; "
            f"color:{GREY}; padding:20px;"
        )
        self.preview_label.setMinimumWidth(420)
        prev_lay.addWidget(self.preview_label, stretch=1)
        splitter.addWidget(preview_container)

        splitter.setSizes([380, 600])
        outer.addWidget(splitter, stretch=1)

        self.start_in.valueChanged.connect(self.update_end)
        self.qty_in.valueChanged.connect(self.update_end)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(400)  # debounce: wait 400ms after the last
        # keystroke / spinbox change before re-rendering the preview so we
        # don't queue a heavy image render on every character typed.
        self.timer.timeout.connect(self.update_preview)
        for w in (self.name_in, self.phone_in, self.addr_in):
            w.textChanged.connect(self.timer.start)
        self.start_in.valueChanged.connect(self.timer.start)
        self.qty_in.valueChanged.connect(self.timer.start)
        self.update_end()

    def update_end(self):
        s = self.start_in.value()
        q = self.qty_in.value()
        e = s + q - 1
        if s == e:
            self.end_label.setText(f"{s:04d}")
        else:
            self.end_label.setText(f"{e:04d}")

    def update_preview(self):
        if self.preview_worker and self.preview_worker.isRunning():
            return
        start = self.start_in.value()
        qty = self.qty_in.value()
        end = start + qty - 1
        if end > rohit.MAX_COUPON:
            self.preview_label.setText(
                f"End {end:04d} exceeds max {rohit.MAX_COUPON}"
            )
            return
        name = self.name_in.text().strip()
        phone = self.phone_in.text().strip()
        addr = self.addr_in.text().strip()
        amount = qty * rohit.PRICE_PER_COUPON
        self.preview_label.setText("Generating preview...")
        self.preview_worker = PreviewWorker(start, end, name, phone, addr, amount=amount)
        self.preview_worker.done.connect(self.show_preview)
        self.preview_worker.start()

    def show_preview(self, qimg):
        if qimg is None:
            self.preview_label.setText("Preview unavailable")
            return
        pix = QPixmap.fromImage(qimg)
        scaled = pix.scaledToWidth(
            max(200, self.preview_label.width() - 20),
            Qt.TransformationMode.SmoothTransformation,
        )
        self.preview_label.setPixmap(scaled)

    def do_sale(self):
        start = self.start_in.value()
        qty = self.qty_in.value()
        end = start + qty - 1
        if end > rohit.MAX_COUPON:
            QMessageBox.warning(
                self, "Invalid", f"End {end:04d} exceeds max {rohit.MAX_COUPON}."
            )
            return
        name = self.name_in.text().strip()
        phone = self.phone_in.text().strip()
        addr = self.addr_in.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing", "Buyer name is required.")
            return
        # The overlap check happens inside SaleWorker (off the GUI thread) so
        # a slow Excel / OneDrive read cannot freeze the window.
        self.sell_btn.setEnabled(False)
        self.status_label.setText("Processing...")
        self.worker = SaleWorker(name, phone, addr, start, qty)
        self.worker.done.connect(self.on_sale_done)
        self.worker.start()

    def on_sale_done(self, msg, success):
        self.sell_btn.setEnabled(True)
        self.status_label.setText(msg)
        if success:
            QMessageBox.information(self, "Success", msg)
            self.parent_window.refresh_all()
        else:
            QMessageBox.warning(self, "Problem", msg)


class PhysicalTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.worker = None
        self.build()

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("PHYSICAL SET GENERATION")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        info = QLabel(
            "Generate coupon strips with no buyer details. Numbers are locked "
            "for physical sale. QR shows only the coupon number(s)."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color:{GREY};")
        lay.addWidget(info)

        form = QFormLayout()
        form.setSpacing(10)
        self.size_combo = QComboBox()
        self.size_combo.addItem(f"Set of 10  — \u20B9{rohit.price_for_set_size(10)}", 10)
        self.size_combo.addItem(f"Set of 5  — \u20B9{rohit.price_for_set_size(5)}", 5)
        self.size_combo.addItem(f"Set of 1  — \u20B9{rohit.price_for_set_size(1)}", 1)
        self.start_in = QSpinBox()
        self.start_in.setRange(1, rohit.MAX_COUPON)
        self.start_in.setValue(1)
        self.sets_in = QSpinBox()
        self.sets_in.setRange(1, 5000)
        self.sets_in.setValue(1)
        form.addRow("Set size:", self.size_combo)
        form.addRow("Start No:", self.start_in)
        form.addRow("Number of sets:", self.sets_in)
        lay.addLayout(form)

        self.plan_label = QLabel("")
        self.plan_label.setStyleSheet(
            f"background:{WHITE}; border:2px solid {GOLD}; border-radius:6px; "
            f"padding:10px; color:{MAROON}; font-weight:bold;"
        )
        self.plan_label.setWordWrap(True)
        lay.addWidget(self.plan_label)

        self.size_combo.currentIndexChanged.connect(self.update_plan)
        self.start_in.valueChanged.connect(self.update_plan)
        self.sets_in.valueChanged.connect(self.update_plan)
        self.update_plan()

        self.progress = QProgressBar()
        self.progress.setValue(0)
        lay.addWidget(self.progress)

        self.gen_btn = QPushButton("Generate Physical Sets")
        self.gen_btn.setMinimumHeight(46)
        self.gen_btn.clicked.connect(self.generate)
        lay.addWidget(self.gen_btn)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet(f"color:{MAROON};")
        lay.addWidget(self.status_label)

        lay.addStretch()

    def update_plan(self):
        size = self.size_combo.currentData()
        start = self.start_in.value()
        num = self.sets_in.value()
        total = size * num
        end = start + total - 1
        set_price = rohit.price_for_set_size(size)
        total_cost = num * set_price
        self.plan_label.setText(
            f"Plan: {num} set(s) of {size} = {total} coupons, "
            f"numbering {start:04d} - {end:04d}  |  "
            f"\u20B9{set_price} x {num} = \u20B9{total_cost}"
            + (f"  (exceeds max {rohit.MAX_COUPON}!)" if end > rohit.MAX_COUPON else "")
        )

    def generate(self):
        size = self.size_combo.currentData()
        start = self.start_in.value()
        num = self.sets_in.value()
        end = start + num * size - 1
        if end > rohit.MAX_COUPON:
            QMessageBox.warning(
                self, "Invalid", f"Would end at {end:04d}, exceeds {rohit.MAX_COUPON}."
            )
            return
        confirm = QMessageBox.question(
            self, "Confirm",
            f"Generate {num} set(s) of {size} = {size*num} coupons "
            f"({start:04d}-{end:04d})?\n"
            f"Donation: \u20B9{rohit.price_for_set_size(size)} x {num} = "
            f"\u20B9{num * rohit.price_for_set_size(size)}"
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.gen_btn.setEnabled(False)
        self.progress.setValue(0)
        self.status_label.setText("Generating...")
        self.worker = PhysicalWorker(size, start, num)
        self.worker.progress.connect(self.on_progress)
        self.worker.done.connect(self.on_done)
        self.worker.start()

    def on_progress(self, cur, total, msg):
        pct = int((cur / total) * 100) if total else 0
        self.progress.setValue(pct)
        self.status_label.setText(msg)

    def on_done(self, msg, success):
        self.progress.setValue(100 if success else 0)
        self.status_label.setText(msg)
        self.gen_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Done", msg)
            self.parent_window.refresh_all()
        else:
            QMessageBox.warning(self, "Problem", msg)


class SalesTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.all_rows = []
        self.build()

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("ALL SALES")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Search:"))
        self.search_in = QLineEdit()
        self.search_in.setPlaceholderText(
            "Filter by name, phone, coupon number, date..."
        )
        self.search_in.textChanged.connect(self.filter_rows)
        search_row.addWidget(self.search_in, stretch=1)
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setProperty("flat", True)
        self.refresh_btn.clicked.connect(self.refresh)
        search_row.addWidget(self.refresh_btn)
        lay.addLayout(search_row)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels([
            "S.No", "Name", "Phone", "Address", "Start", "End", "Qty", "Date", "Type", "Donation"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        # Enable clickable-header sorting.  Numeric columns use
        # NumericTableWidgetItem so Qt compares them as ints, not as strings.
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for i in [0, 2, 4, 5, 6, 7, 8, 9]:
            self.table.horizontalHeader().setSectionResizeMode(
                i, QHeaderView.ResizeMode.ResizeToContents
            )
        lay.addWidget(self.table, stretch=1)

        btn_row = QHBoxLayout()
        self.view_btn = QPushButton("View Coupon PNG")
        self.view_btn.setProperty("flat", True)
        self.view_btn.clicked.connect(self.view_png)
        btn_row.addWidget(self.view_btn)
        self.edit_btn = QPushButton("Edit Selected Sale")
        self.edit_btn.setProperty("flat", True)
        self.edit_btn.clicked.connect(self.edit_selected)
        btn_row.addWidget(self.edit_btn)
        self.delete_btn = QPushButton("Delete Selected Sale")
        self.delete_btn.clicked.connect(self.delete_selected)
        btn_row.addWidget(self.delete_btn)
        lay.addLayout(btn_row)

    def refresh(self):
        # Disk I/O happens in the main window's background refresh worker;
        # this tab just re-filters the last snapshot it received.
        self.parent_window.request_refresh()

    def apply_snapshot(self, rows, ranges):
        self.all_rows = list(rows)
        self.filter_rows()

    def filter_rows(self):
        q = self.search_in.text().strip().lower()
        shown = []
        for r in self.all_rows:
            row_str = " ".join(str(c or "") for c in r).lower()
            if not q or q in row_str:
                shown.append(r)

        # Sorting is enabled on the table, so building rows triggers Qt's
        # row-comparison.  Temporarily disable sorting while we populate,
        # then re-apply the default sort (S.No descending = newest first).
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(shown))
        for i, r in enumerate(shown):
            # Donation amount (index 10 in the row tuple).
            if len(r) >= 11 and r[10] is not None:
                amt_val = r[10]
                amt = f"\u20B9{amt_val}"
            else:
                qty = (r[6] if len(r) >= 7 else 0) or 0
                stype = r[8] if len(r) >= 9 else "SALE"
                amt_val = (rohit.price_for_set_size(qty) if stype == "PHYSICAL"
                           else qty * rohit.PRICE_PER_COUPON)
                amt = f"\u20B9{amt_val}"

            sno = r[0] if r[0] is not None else ""
            start = r[4] if r[4] is not None else ""
            end = r[5] if r[5] is not None else ""
            qty_v = r[6] if r[6] is not None else ""

            # Numeric columns -> NumericTableWidgetItem for correct sorting.
            sno_item = NumericTableWidgetItem(sno)
            sno_item.setText(str(sno))
            start_item = NumericTableWidgetItem(start)
            start_item.setText(str(start))
            end_item = NumericTableWidgetItem(end)
            end_item.setText(str(end))
            qty_item = NumericTableWidgetItem(qty_v)
            qty_item.setText(str(qty_v))
            amt_item = NumericTableWidgetItem(amt_val)
            amt_item.setText(amt)

            name_disp = str(r[1] or "") if r[1] else "<PHYSICAL>"
            items = [
                sno_item,
                QTableWidgetItem(name_disp),
                QTableWidgetItem(str(r[2] or "")),
                QTableWidgetItem(str(r[3] or "")),
                start_item,
                end_item,
                qty_item,
                QTableWidgetItem(str(r[7] or "")),
                QTableWidgetItem(str(r[8] if len(r) >= 9 else "SALE")),
                amt_item,
            ]
            for c, v in enumerate(items):
                self.table.setItem(i, c, v)

        self.table.resizeColumnsToContents()
        # Re-enable sorting and apply the default sort: S.No (column 0)
        # descending, which shows the newest sale first — preserving the
        # previous "recentest first" behaviour until the user clicks a header.
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(0, Qt.SortOrder.DescendingOrder)

    def view_png(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Select", "Select a row first.")
            return
        start_item = self.table.item(row, 4)
        end_item = self.table.item(row, 5)
        name_item = self.table.item(row, 1)
        if not start_item or not end_item:
            return
        start = int(start_item.text())
        end = int(end_item.text())
        is_range = start != end
        num_part = f"{start:04d}-{end:04d}" if is_range else f"{start:04d}"
        name = name_item.text() if name_item else ""
        if name and name != "<PHYSICAL>":
            invalid = '<>:"/\\|?*'
            safe = "".join("_" if c in invalid else c for c in name).strip()
            safe = safe.replace(" ", "_")
            filename = f"{safe}_{num_part}.png"
        else:
            filename = f"coupon_{num_part}.png"
        path = rohit.OUTPUT_DIR / filename
        if not path.exists():
            QMessageBox.information(self, "Not Found", f"{filename} not found.")
            return
        dlg = ImageViewerDialog(path, f"Coupon {num_part}")
        dlg.exec()

    def delete_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Select", "Select a row first.")
            return
        sno_item = self.table.item(row, 0)
        if not sno_item:
            return
        sno = int(sno_item.text())
        name_item = self.table.item(row, 1)
        name_disp = name_item.text() if name_item else ""
        confirm = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete sale #{sno} ({name_disp})?\nCoupon numbers will be freed."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        del_png = QMessageBox.question(
            self, "Delete PNG", "Also delete the coupon PNG file?"
        ) == QMessageBox.StandardButton.Yes
        self._start_delete("sale", (sno, del_png), f"Sale #{sno} deleted.")

    def edit_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Select", "Select a row first.")
            return
        sno_item = self.table.item(row, 0)
        if not sno_item:
            return
        try:
            sno = int(sno_item.text())
        except (ValueError, TypeError):
            return

        # Read the currently-displayed values so the dialog starts with them.
        name_item = self.table.item(row, 1)
        phone_item = self.table.item(row, 2)
        addr_item = self.table.item(row, 3)
        type_item = self.table.item(row, 8)

        name = name_item.text() if name_item else ""
        phone = phone_item.text() if phone_item else ""
        address = addr_item.text() if addr_item else ""
        stype = type_item.text() if type_item else "SALE"
        if name == "<PHYSICAL>":
            name = ""  # physical rows start with blank buyer fields

        dlg = EditSaleDialog(sno, name=name, phone=phone, address=address,
                             stype=stype, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        new_name, new_phone, new_addr = dlg.values()
        self._edit_ok_msg = f"Sale #{sno} updated."
        self._edit_token = ("edit", sno)
        self.edit_btn.setEnabled(False)
        self._update_worker = UpdateWorker(
            sno, new_name, new_phone, new_addr,
            caller_token=self._edit_token,
        )
        self._update_worker.done.connect(self._on_update_done)
        self._update_worker.start()

    def _on_update_done(self, token, success, msg):
        if token != getattr(self, "_edit_token", None):
            return
        self.edit_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Updated", self._edit_ok_msg)
            self.parent_window.refresh_all()
        else:
            QMessageBox.warning(self, "Problem", msg)

    def _start_delete(self, token, args, ok_msg):
        # Disable interaction while the background delete runs so the UI
        # stays responsive without letting the user fire a second delete.
        self._delete_ok_msg = ok_msg
        self._delete_token = token
        if token == "sale":
            self._delete_worker = DeleteWorker(
                rohit.delete_sale, *args, caller_token=token
            )
        else:
            return
        self._delete_worker.done.connect(self._on_delete_done)
        self._delete_worker.start()

    def _on_delete_done(self, token, success, msg):
        if token != getattr(self, "_delete_token", None):
            return
        if success:
            QMessageBox.information(self, "Deleted", self._delete_ok_msg)
            self.parent_window.refresh_all()
        else:
            QMessageBox.warning(self, "Not Found", msg)


class ManageTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.build()

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("MANAGE SALES")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        g1 = QGroupBox("Delete Single Sale")
        g1_lay = QFormLayout(g1)
        self.del_sno = QSpinBox()
        self.del_sno.setRange(1, 999999)
        self.del_sno.setValue(1)
        g1_lay.addRow("S.No:", self.del_sno)
        self.del_png_chk = QCheckBox("Also delete PNG file")
        self.del_png_chk.setChecked(True)
        g1_lay.addRow(self.del_png_chk)
        b1 = QPushButton("Delete Sale")
        b1.clicked.connect(self.delete_sale)
        g1_lay.addRow(b1)
        lay.addWidget(g1)

        g2 = QGroupBox("Delete Physical Sets by Range")
        g2_lay = QFormLayout(g2)
        self.rng_start = QSpinBox()
        self.rng_start.setRange(1, rohit.MAX_COUPON)
        self.rng_start.setValue(1)
        self.rng_end = QSpinBox()
        self.rng_end.setRange(1, rohit.MAX_COUPON)
        self.rng_end.setValue(100)
        g2_lay.addRow("Range Start:", self.rng_start)
        g2_lay.addRow("Range End:", self.rng_end)
        self.rng_png_chk = QCheckBox("Also delete PNG files")
        self.rng_png_chk.setChecked(True)
        g2_lay.addRow(self.rng_png_chk)
        b2 = QPushButton("Delete Physical in Range")
        b2.clicked.connect(self.delete_physical_range)
        g2_lay.addRow(b2)
        lay.addWidget(g2)

        g3 = QGroupBox("Bulk Deletes")
        g3_lay = QVBoxLayout(g3)
        b3 = QPushButton("Delete ALL Physical Sets")
        b3.clicked.connect(self.delete_all_physical)
        g3_lay.addWidget(b3)
        b4 = QPushButton("Delete ALL Sales (physical + normal)")
        b4.setStyleSheet(f"background:{MAROON}; color:{GOLD_LIGHT};")
        b4.clicked.connect(self.delete_all_sales)
        g3_lay.addWidget(b4)
        lay.addWidget(g3)

        lay.addStretch()

    def delete_sale(self):
        sno = self.del_sno.value()
        del_png = self.del_png_chk.isChecked()
        confirm = QMessageBox.question(
            self, "Confirm", f"Delete sale #{sno}?"
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._start_delete("sale", (rohit.delete_sale, sno, del_png),
                           f"Sale #{sno} deleted.")

    def delete_physical_range(self):
        s = self.rng_start.value()
        e = self.rng_end.value()
        if e < s:
            QMessageBox.warning(self, "Invalid", "End must be >= start.")
            return
        confirm = QMessageBox.question(
            self, "Confirm",
            f"Delete all physical sets within {s:04d}-{e:04d}?\n"
            "Normal sales are kept."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._start_delete("phys_range",
                           (rohit.delete_physical_by_range, s, e,
                            self.rng_png_chk.isChecked()),
                           f"Physical sets in {s:04d}-{e:04d} deleted.")

    def delete_all_physical(self):
        confirm = QMessageBox.question(
            self, "Confirm",
            "Delete ALL physical sets? Normal sales are kept. This cannot be undone."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._start_delete("phys_all", (rohit.delete_all_physical, True),
                           "All physical sets deleted.")

    def delete_all_sales(self):
        confirm = QMessageBox.question(
            self, "Confirm",
            "Delete ALL sales (physical + normal)? This CANNOT be undone."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        png_confirm = QMessageBox.question(
            self, "Delete PNGs", "Also delete all generated PNG files?"
        ) == QMessageBox.StandardButton.Yes
        self._start_delete("all_sales",
                           (rohit.delete_all_sales, png_confirm),
                           "All sales deleted.")

    def _start_delete(self, token, fn_args, ok_msg):
        # Run the delete in a background thread so the UI keeps responding
        # while Excel is saved and PNG files are removed.
        self._delete_ok_msg = ok_msg
        self._delete_token = token
        fn = fn_args[0]
        args = fn_args[1:]
        self._delete_worker = DeleteWorker(fn, *args, caller_token=token)
        self._delete_worker.done.connect(self._on_delete_done)
        self._delete_worker.start()

    def _on_delete_done(self, token, success, msg):
        if token != getattr(self, "_delete_token", None):
            return
        if success:
            QMessageBox.information(self, "Done", self._delete_ok_msg)
            self.parent_window.refresh_all()
        else:
            QMessageBox.warning(self, "Problem", msg)


class GapsTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.build()

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("COUPON NUMBER GAPS")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        self.summary = QLabel("")
        self.summary.setStyleSheet(
            f"background:{WHITE}; border:2px solid {GOLD}; border-radius:6px; "
            f"padding:10px; color:{MAROON}; font-weight:bold;"
        )
        self.summary.setWordWrap(True)
        lay.addWidget(self.summary)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Gap Range", "Count"])
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        lay.addWidget(self.table, stretch=1)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setProperty("flat", True)
        refresh_btn.clicked.connect(self.refresh)
        lay.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignRight)

    def refresh(self):
        self.parent_window.request_refresh()

    def apply_snapshot(self, rows, ranges):
        if not ranges:
            self.summary.setText("No coupons assigned yet.")
            self.table.setRowCount(0)
            return
        assigned = set()
        for s, e in ranges:
            for n in range(s, e + 1):
                assigned.add(n)
        low, high = min(assigned), max(assigned)
        gaps = [n for n in range(low, high + 1) if n not in assigned]

        self.summary.setText(
            f"Lowest assigned: {low:04d}  |  Highest assigned: {high:04d}  |  "
            f"Total in span: {high-low+1}  |  Assigned: {len(assigned)}  |  "
            f"Gaps: {len(gaps)}"
        )

        if not gaps:
            self.table.setRowCount(1)
            self.table.setItem(0, 0, QTableWidgetItem(
                f"No gaps. Every number from {low:04d} to {high:04d} is assigned."
            ))
            self.table.setItem(0, 1, QTableWidgetItem("0"))
            return

        collapsed = []
        run_s = gaps[0]
        run_e = gaps[0]
        for n in gaps[1:]:
            if n == run_e + 1:
                run_e = n
            else:
                collapsed.append((run_s, run_e))
                run_s = n
                run_e = n
        collapsed.append((run_s, run_e))

        self.table.setRowCount(len(collapsed))
        for i, (s, e) in enumerate(collapsed):
            if s == e:
                label = f"{s:04d}"
                count = "1"
            else:
                label = f"{s:04d}-{e:04d}"
                count = str(e - s + 1)
            self.table.setItem(i, 0, QTableWidgetItem(label))
            self.table.setItem(i, 1, QTableWidgetItem(count))


class LuckyDrawTab(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self._worker = None
        self.build()

    def build(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(18, 18, 18, 18)

        title = QLabel("LUCKY DRAW")
        title.setProperty("heading", True)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)

        info = QLabel(
            "Draws 50 distinct winning coupon numbers from ALL 9999 coupons "
            "(1-9999, sold or unsold - 1-in-9999 chance each) - 10 each for "
            "Consolation D, C, B, A (40 consolation) plus 10 main prizes "
            "(1st-10th). Unsold winning numbers are marked UNSOLD. Results "
            "are saved to the workbook and cannot be re-run until cleared."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color:{GREY};")
        lay.addWidget(info)

        self.status_label = QLabel("Checking draw status...")
        self.status_label.setStyleSheet(
            f"background:{WHITE}; border:2px solid {GOLD}; border-radius:6px; "
            f"padding:10px; color:{MAROON}; font-weight:bold;"
        )
        self.status_label.setWordWrap(True)
        lay.addWidget(self.status_label)

        btn_row = QHBoxLayout()
        self.draw_btn = QPushButton("Conduct Draw")
        self.draw_btn.setMinimumHeight(46)
        self.draw_btn.clicked.connect(self.conduct_draw)
        btn_row.addWidget(self.draw_btn)
        self.clear_btn = QPushButton("Clear Results")
        self.clear_btn.setProperty("flat", True)
        self.clear_btn.setMinimumHeight(46)
        self.clear_btn.clicked.connect(self.clear_results)
        self.clear_btn.setEnabled(False)
        btn_row.addWidget(self.clear_btn)
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setProperty("flat", True)
        refresh_btn.clicked.connect(self.load_results)
        btn_row.addWidget(refresh_btn)
        lay.addLayout(btn_row)

        winners_title = QLabel("WINNERS")
        winners_title.setProperty("heading", True)
        lay.addWidget(winners_title)

        self.table = QTableWidget(0, 12)
        self.table.setHorizontalHeaderLabels([
            "S.No", "Prize", "Gift Item", "Coupon No", "Set Range", "Buyer",
            "Phone", "Address", "Qty", "Date Sold", "Type", "Category"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for i in [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]:
            hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.table, stretch=1)

        self.msg_label = QLabel("")
        self.msg_label.setWordWrap(True)
        self.msg_label.setStyleSheet(f"color:{MAROON}; font-weight:bold;")
        lay.addWidget(self.msg_label)

    def on_show(self):
        self.load_results()

    def _start_worker(self, action):
        if self._worker and self._worker.isRunning():
            return
        self._worker = DrawActionWorker(action)
        self._worker.done.connect(self.on_worker_done)
        self._worker.start()

    def load_results(self):
        self.status_label.setText("Checking draw status...")
        self.draw_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self._start_worker("load")

    def conduct_draw(self):
        confirm = QMessageBox.question(
            self, "Confirm Draw",
            "Conduct the Lucky Draw now?\n"
            "50 distinct winners will be randomly picked from ALL 9999 coupons "
            "(1-in-9999 chance each).\n"
            "Unsold winning numbers will be marked UNSOLD.\n"
            "This cannot be undone."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.draw_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.msg_label.setText("Drawing 50 winners from all 9999 coupons...")
        self._start_worker("draw")

    def clear_results(self):
        confirm = QMessageBox.question(
            self, "Confirm Clear",
            "Clear the saved Lucky Draw results?\n"
            "A new draw can then be conducted."
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.msg_label.setText("Clearing...")
        self._start_worker("clear")

    def on_worker_done(self, tag, success, msg, results):
        self.msg_label.setText(msg if not success else "")
        if not success:
            QMessageBox.warning(self, "Problem", msg)
            self.load_results()
            return

        if tag == "clear":
            QMessageBox.information(self, "Cleared", msg)
            self.load_results()
            return

        if tag == "draw":
            QMessageBox.information(self, "Draw Complete", msg)

        # tag == "draw" or "load" -> populate the table
        if results:
            drawn_at = results[0].get("drawn_at", "—") if isinstance(results[0], dict) else "—"
            self.status_label.setText(
                f"Draw completed on {drawn_at} - 50 winners saved."
            )
            self.draw_btn.setEnabled(False)
            self.clear_btn.setEnabled(True)
            _cat_order = {"D": 0, "C": 1, "B": 2, "A": 3, "MAIN": 4}
            _main_rank = {
                "1st Prize": 0, "2nd Prize": 1, "3rd Prize": 2, "4th Prize": 3,
                "5th Prize": 4, "6th Prize": 5, "7th Prize": 6, "8th Prize": 7,
                "9th Prize": 8, "10th Prize": 9,
            }
            def _sort_key(r):
                cat = r.get("category") or ("MAIN" if not str(r.get("prize", "")).startswith("Consolation") else "")
                prize = str(r.get("prize", ""))
                if cat == "MAIN":
                    return (4, _main_rank.get(prize, 99))
                return (_cat_order.get(cat, 9), 0)
            ordered = sorted(results, key=_sort_key)
            self.table.setRowCount(len(ordered))
            for i, r in enumerate(ordered):
                coupon = r.get("coupon_no", "")
                coupon_str = f"{int(coupon):04d}" if coupon != "" else "—"
                set_range = r.get("set_range", "") or r.get("coupon_range_start")
                if not set_range and r.get("coupon_range_start") is not None:
                    s = r.get("coupon_range_start")
                    e = r.get("coupon_range_end")
                    set_range = f"{int(s):04d}-{int(e):04d}" if s != e else ""
                else:
                    set_range = r.get("set_range", "") or ""
                cat = r.get("category", "") or ""
                if not cat:
                    cat = "MAIN" if not str(r.get("prize", "")).startswith("Consolation") else ""
                cat_label = ("Main" if cat == "MAIN"
                             else (f"Consolation {cat}" if cat else "—"))
                vals = [
                    str(i + 1),
                    str(r.get("prize", "")),
                    str(r.get("gift", "")),
                    coupon_str,
                    str(set_range) if set_range else "—",
                    str(r.get("buyer", "")),
                    str(r.get("phone", "") or "—"),
                    str(r.get("address", "") or "—"),
                    str(r.get("qty", "") or "—"),
                    str(r.get("date_sold", "") or "—"),
                    str(r.get("type", "")),
                    cat_label,
                ]
                is_unsold = (r.get("type") == "UNSOLD" or r.get("buyer") == "UNSOLD")
                for c, v in enumerate(vals):
                    item = QTableWidgetItem(v)
                    if is_unsold:
                        if c == 5:
                            item.setForeground(QColor(MAROON))
                        else:
                            item.setForeground(QColor(GREY))
                    self.table.setItem(i, c, item)
            self.table.resizeColumnsToContents()
        else:
            self.status_label.setText(
                "No draw conducted yet. Click Conduct Draw to pick 50 winners "
                "from all 9999 coupons (40 consolation + 10 main)."
            )
            self.draw_btn.setEnabled(True)
            self.clear_btn.setEnabled(False)
            self.table.setRowCount(0)


# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Jai Bhadra Foundation - Lucky Draw Manager")
        self.resize(1180, 780)

        central = QWidget()
        outer = QVBoxLayout(central)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(10)

        banner = QLabel("JAI BHADRA FOUNDATION  -  LUCKY DRAW COUPON MANAGER")
        banner.setAlignment(Qt.AlignmentFlag.AlignCenter)
        banner.setStyleSheet(
            f"background:{MAROON}; color:{GOLD_LIGHT}; "
            f"padding:14px; font-size:20px; font-weight:bold; "
            f"border:2px solid {GOLD}; border-radius:10px;"
        )
        outer.addWidget(banner)

        self.tabs = QTabWidget()
        self.dashboard = DashboardTab(self)
        self.new_sale = NewSaleTab(self)
        self.physical = PhysicalTab(self)
        self.sales = SalesTab(self)
        self.manage = ManageTab(self)
        self.gaps = GapsTab(self)
        self.lucky_draw = LuckyDrawTab(self)
        self.tabs.addTab(self.dashboard, "Dashboard")
        self.tabs.addTab(self.new_sale, "New Sale")
        self.tabs.addTab(self.physical, "Physical Sets")
        self.tabs.addTab(self.sales, "Sales")
        self.tabs.addTab(self.manage, "Manage")
        self.tabs.addTab(self.gaps, "Gaps")
        self.tabs.addTab(self.lucky_draw, "Lucky Draw")
        self.tabs.currentChanged.connect(self.on_tab_changed)
        outer.addWidget(self.tabs)

        footer = QLabel(
            f"Draw: {rohit.DRAW_LINE}  |  Max coupons: {rohit.MAX_COUPON}  |  "
            f"Website: {rohit.WEBSITE}"
        )
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet(f"color:{GREY}; font-size:12px; padding:6px;")
        outer.addWidget(footer)

        self.setCentralWidget(central)

        # Shared refresh worker.  All tab refreshes route through here so
        # there is only ever ONE background load of coupon_sales.xlsx at a
        # time, and the GUI thread never touches the file directly.  This is
        # what stops the "Not Responding" freeze — the GUI thread no longer
        # waits on OneDrive's file lock.
        self._refresh_worker = None
        self._refresh_pending = False
        self._last_snapshot = ([], [])

        # Defer the first Excel probe + refresh until *after* the window has
        # painted.  Doing it inside __init__ blocks the GUI thread and is what
        # made the app show "Not Responding" on startup, especially when
        # coupon_sales.xlsx lives on OneDrive.
        QTimer.singleShot(0, self._startup_init)

    def _startup_init(self):
        # Runs in a background thread via request_refresh; the file probe
        # itself also happens there, off the GUI thread.
        self.request_refresh()

    def on_tab_changed(self, idx):
        widget = self.tabs.widget(idx)
        if widget is self.lucky_draw:
            self.lucky_draw.on_show()
        elif hasattr(widget, "refresh"):
            widget.refresh()

    def request_refresh(self):
        # Coalesce rapid refresh requests (e.g. several tabs refreshing at
        # once) into a single background load.  If a load is already in
        # flight, just flag that another is wanted when it finishes.
        if self._refresh_worker is not None and self._refresh_worker.isRunning():
            self._refresh_pending = True
            return
        self._refresh_worker = RefreshWorker()
        self._refresh_worker.done.connect(self._on_refresh_done)
        self._refresh_worker.start()

    def _on_refresh_done(self, rows, ranges, error):
        if error is not None:
            QMessageBox.critical(self, "Refresh Error", str(error))
        else:
            self._last_snapshot = (rows, ranges)
            self.dashboard.apply_snapshot(rows, ranges)
            self.sales.apply_snapshot(rows, ranges)
            self.gaps.apply_snapshot(rows, ranges)

        # If another refresh was requested while we were loading, kick it
        # off again now so the data stays fresh.
        self._refresh_worker = None
        if self._refresh_pending:
            self._refresh_pending = False
            self.request_refresh()

    def refresh_all(self):
        # Called after a sale / delete.  Routes through the background worker
        # so the GUI thread never blocks on Excel / OneDrive.
        self.request_refresh()
        if hasattr(self.new_sale, "update_preview"):
            self.new_sale.update_preview()


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    app.setApplicationName("JBF Lucky Draw")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()