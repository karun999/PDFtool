"""
PDF Studio Pro
--------------
A simple desktop tool to Merge, Split, Extract and Organize PDF pages.

Install the libraries once (in Command Prompt):
    pip install PyQt5 pypdf cryptography pymupdf rapidocr_onnxruntime openpyxl

pymupdf draws page previews and is needed for OCR.
rapidocr_onnxruntime is the offline OCR engine (no Tesseract needed).
openpyxl writes the invoice Excel file.
Without them, Merge / Split / Extract / Organize still work.

Then open this file in IDLE and press F5 to run.
"""

import base64
import datetime
import hashlib
import hmac
import html
import json
import importlib.util
import os
import re
import sys
import threading
import time
from collections import Counter

# Load the OCR runtime BEFORE Qt. PyQt5 ships an older Microsoft C++ runtime
# DLL (msvcp140.dll); if Qt loads it first, onnxruntime can crash on Windows.
if importlib.util.find_spec("onnxruntime") is not None:
    try:
        import onnxruntime  # noqa: F401
    except Exception:
        pass

from PyQt5.QtCore import QPointF, QRect, QRectF, QSize, Qt, QThread, QTimer, QUrl, pyqtSignal
from PyQt5.QtGui import (
    QColor, QDesktopServices, QDrag, QFont, QIcon, QImage, QPainter, QPen,
    QPixmap, QTransform,
)
from PyQt5.QtWidgets import (
    QAbstractItemView, QApplication, QButtonGroup, QCheckBox, QFileDialog,
    QFrame, QGraphicsDropShadowEffect, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QListView, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QScrollArea,
    QSizePolicy, QSlider, QSpinBox, QStackedWidget, QTableWidget,
    QTableWidgetItem, QTabWidget, QTextBrowser, QVBoxLayout, QWidget,
    QHeaderView,
)
from pypdf import PdfReader, PdfWriter

try:
    import pymupdf as fitz  # PyMuPDF - used only for page preview images
    HAS_PREVIEWS = True
except ImportError:
    try:
        import fitz  # older PyMuPDF versions
        HAS_PREVIEWS = True
    except ImportError:
        HAS_PREVIEWS = False

# OCR engine is loaded only when OCR is used (it takes a moment to start)
HAS_OCR = HAS_PREVIEWS and importlib.util.find_spec("rapidocr_onnxruntime") is not None
HAS_XLSX = importlib.util.find_spec("openpyxl") is not None

APP_NAME = "PDF Studio Pro"
APP_VERSION = "2.0"

# Colour themes for each tool (gradient start, gradient end)
MERGE_COLORS = ("#ff512f", "#dd2476")
SPLIT_COLORS = ("#0f9b8e", "#2bc36a")
EXTRACT_COLORS = ("#4776e6", "#8e54e9")
ORGANIZE_COLORS = ("#ff9a00", "#ff5e00")
OCR_COLORS = ("#b24592", "#f15f79")
SUMMARY_COLORS = ("#1d976c", "#2bb3a3")
INVOICE_COLORS = ("#3a7bd5", "#00b4db")
REDACT_COLORS = ("#232526", "#d31027")

# Action colours (Redaction tool and others)
BLUE = ("#2196f3", "#1565c0")      # main actions
GREEN = ("#43a047", "#2e7d32")     # success / save / apply
ORANGE = ("#ff9800", "#f57c00")    # important actions
RED = ("#ef5350", "#c62828")       # destructive actions
PURPLE = ("#ab47bc", "#6a1b9a")    # secondary actions

# Each output file in Organize & Split gets its own border colour
SEGMENT_COLORS = ["#ff6d00", "#7c4dff", "#00bfa5", "#f50057",
                  "#2979ff", "#43a047", "#ffab00", "#d500f9"]

# Size of one page tile in the Organize & Split grid
TILE_W, TILE_H = 136, 186

# Data stored on each page tile
ROLE_PAGE = Qt.UserRole        # original page index (0-based)
ROLE_ROT = Qt.UserRole + 1     # extra rotation in degrees
ROLE_BREAK = Qt.UserRole + 2   # True = split after this page


# ======================================================================
#  Trial and daily usage limits
# ======================================================================
#  - The app works for TRIAL_DAYS days, counted from the first time it is
#    opened on a PC. After that it refuses to start.
#  - It can be used DAILY_LIMIT_MINUTES minutes per calendar day. When the
#    limit is reached it warns the user and closes; it opens again tomorrow.
#  - Limits apply to the .exe only (so you can keep developing in IDLE).
#    Set ENFORCE_LIMITS_ONLY_IN_EXE = False to test them in IDLE.

TRIAL_DAYS = 20
DAILY_LIMIT_MINUTES = 120
WARNING_MINUTES = (10, 1)          # warn when this many minutes are left today
ENFORCE_LIMITS_ONLY_IN_EXE = True
CLOCK_TOLERANCE_SECONDS = 600      # allow small automatic clock corrections

_GUARD_SECRET = b"PDFStudioPro::trial::7f3c9a51e2b84d06"
_GUARD_FILE = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"),
                           "PDFStudioPro", ".session.dat")
_GUARD_REG_KEY = r"Software\PDFStudioPro"
_GUARD_REG_VALUE = "SessionState"

try:
    import winreg
except ImportError:          # not Windows
    winreg = None


def _machine_id():
    if winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r"SOFTWARE\Microsoft\Cryptography") as key:
                return winreg.QueryValueEx(key, "MachineGuid")[0]
        except OSError:
            pass
    return os.environ.get("COMPUTERNAME", "") or os.path.expanduser("~")


class UsageGuard:
    """Keeps first-use date and minutes used per day, signed so that edits are
    detected, stored in two places (a hidden file and the Windows registry)."""

    def __init__(self, enforce=None):
        if enforce is None:
            enforce = getattr(sys, "frozen", False) or not ENFORCE_LIMITS_ONLY_IN_EXE
        self.enforce = enforce
        self.key = hashlib.sha256(_GUARD_SECRET + _machine_id().encode()).digest()
        self.state = None
        self.warned = set()
        self._last_mono = self._mono()

    # these three are separate so they can be tested
    def _now(self):
        return time.time()

    def _today(self):
        return datetime.date.today()

    def _mono(self):
        return time.monotonic()

    # ---- signed storage ----
    def _encode(self, state):
        payload = base64.b64encode(json.dumps(state, sort_keys=True).encode()).decode()
        sig = hmac.new(self.key, payload.encode(), hashlib.sha256).hexdigest()
        return payload + "." + sig

    def _decode(self, blob):
        """Returns (state or None, tampered)."""
        if not blob:
            return None, False
        try:
            payload, sig = blob.strip().rsplit(".", 1)
            good = hmac.new(self.key, payload.encode(), hashlib.sha256).hexdigest()
            if not hmac.compare_digest(sig, good):
                return None, True
            return json.loads(base64.b64decode(payload)), False
        except Exception:
            return None, True

    def _read_blobs(self):
        blobs = []
        try:
            with open(_GUARD_FILE, "r", encoding="ascii") as handle:
                blobs.append(handle.read())
        except OSError:
            pass
        if winreg is not None:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _GUARD_REG_KEY) as key:
                    blobs.append(winreg.QueryValueEx(key, _GUARD_REG_VALUE)[0])
            except OSError:
                pass
        return blobs

    def load(self):
        merged, tampered = None, False
        for blob in self._read_blobs():
            state, bad = self._decode(blob)
            tampered = tampered or bad
            if state is None:
                continue
            if merged is None:
                merged = state
                continue
            # keep the most restrictive values from both copies
            merged["first_run"] = min(merged["first_run"], state["first_run"])
            merged["last_seen"] = max(merged["last_seen"], state["last_seen"])
            merged["locked"] = merged.get("locked") or state.get("locked") or ""
            for day, secs in state.get("usage", {}).items():
                merged["usage"][day] = max(merged["usage"].get(day, 0), secs)
        if tampered and merged is not None:
            merged["locked"] = merged.get("locked") or "tamper"
        elif tampered:
            merged = {"first_run": self._today().isoformat(), "last_seen": self._now(),
                      "usage": {}, "locked": "tamper"}
        return merged

    def save(self, state):
        cutoff = (self._today() - datetime.timedelta(days=60)).isoformat()
        state["usage"] = {d: s for d, s in state["usage"].items() if d >= cutoff}
        blob = self._encode(state)
        try:
            os.makedirs(os.path.dirname(_GUARD_FILE), exist_ok=True)
            with open(_GUARD_FILE, "w", encoding="ascii") as handle:
                handle.write(blob)
        except OSError:
            pass
        if winreg is not None:
            try:
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _GUARD_REG_KEY) as key:
                    winreg.SetValueEx(key, _GUARD_REG_VALUE, 0, winreg.REG_SZ, blob)
            except OSError:
                pass

    # ---- rules ----
    def _check(self, state):
        """Updates locks; returns a message if the app must not run."""
        now, today = self._now(), self._today()
        first = datetime.date.fromisoformat(state["first_run"])
        if now < state["last_seen"] - CLOCK_TOLERANCE_SECONDS or today < first:
            state["locked"] = state.get("locked") or "clock"
        state["last_seen"] = max(state["last_seen"], now)
        if (today - first).days >= TRIAL_DAYS:
            state["locked"] = state.get("locked") or "expired"

        reason = state.get("locked")
        if reason == "expired":
            last_day = first + datetime.timedelta(days=TRIAL_DAYS - 1)
            return (f"Your {TRIAL_DAYS}-day trial of {APP_NAME} has ended "
                    f"(the last day was {last_day:%d %b %Y}).\n\n"
                    "The application can no longer be used. "
                    "Please contact your provider for the full version.")
        if reason:
            return (f"{APP_NAME} has been locked because its licence information was "
                    "changed or the computer clock was set back.\n\n"
                    "Please contact your provider.")
        if self.used_today(state) >= DAILY_LIMIT_MINUTES * 60:
            return (f"You have used {APP_NAME} for {DAILY_LIMIT_MINUTES} minutes today, "
                    "which is the daily limit.\n\nIt will be available again tomorrow.")
        return None

    def used_today(self, state=None):
        state = state or self.state
        return state["usage"].get(self._today().isoformat(), 0) if state else 0

    def days_left(self):
        if not self.state:
            return TRIAL_DAYS
        first = datetime.date.fromisoformat(self.state["first_run"])
        return max(0, TRIAL_DAYS - (self._today() - first).days)

    def minutes_left_today(self):
        return max(0, DAILY_LIMIT_MINUTES - int(self.used_today() // 60))

    def start_check(self):
        """Call once before the window opens. Returns a message if blocked."""
        if not self.enforce:
            return None
        state = self.load()
        if state is None:     # first run on this PC
            state = {"first_run": self._today().isoformat(), "last_seen": self._now(),
                     "usage": {}, "locked": ""}
        message = self._check(state)
        self.save(state)
        self.state = state
        self._last_mono = self._mono()
        return message

    def tick(self):
        """Call regularly while the app is open. Adds the time used since the
        last call. Returns (stop_message, warning_message)."""
        if not self.enforce:
            return None, None
        mono = self._mono()
        elapsed = max(0.0, mono - self._last_mono)   # immune to clock changes
        self._last_mono = mono
        state = self.load() or self.state            # re-read: other windows may add time
        day = self._today().isoformat()
        state["usage"][day] = state["usage"].get(day, 0) + elapsed
        message = self._check(state)
        self.save(state)
        self.state = state
        if message:
            return message, None
        left = DAILY_LIMIT_MINUTES * 60 - self.used_today()
        for minutes in sorted(WARNING_MINUTES):
            if left <= minutes * 60 and minutes not in self.warned:
                self.warned.update(m for m in WARNING_MINUTES if m >= minutes)
                return None, (f"Only {max(1, round(left / 60))} minute(s) of today's "
                              f"{DAILY_LIMIT_MINUTES}-minute limit are left.\n\n"
                              "Please save your work - the app will close automatically.")
        return None, None

    def summary(self):
        if not self.enforce:
            return ""
        return (f"Trial: {self.days_left()} day(s) left\n"
                f"Today: {self.minutes_left_today()} of {DAILY_LIMIT_MINUTES} min left")


# ======================================================================
#  PDF helper functions (no user interface code here)
# ======================================================================

def parse_page_ranges(text, total_pages):
    """
    Turn text like '1, 3-5 and 8-10' into [(1, 1), (3, 5), (8, 10)].
    Pages are 1-based and ranges are inclusive.
    Raises ValueError with a friendly message if the input is wrong.
    """
    cleaned = text.strip()
    if not cleaned:
        raise ValueError("Please enter the pages, for example: 1, 3-5, 8-10")

    # Accept 'and', ';' as separators and '-', '–' or 'to' inside a range
    cleaned = re.sub(r"\band\b", ",", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace(";", ",")
    cleaned = re.sub(r"\s*(?:-|–|\bto\b)\s*", "-", cleaned, flags=re.IGNORECASE)
    parts = [p for p in re.split(r"[,\s]+", cleaned) if p]

    ranges = []
    for part in parts:
        match = re.fullmatch(r"(\d+)-(\d+)", part)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
        elif part.isdigit():
            start = end = int(part)
        else:
            raise ValueError(f"'{part}' is not a valid page number or range.")

        if start < 1:
            raise ValueError("Page numbers start from 1.")
        if start > end:
            raise ValueError(f"Range '{start}-{end}' is backwards. Use '{end}-{start}'.")
        if end > total_pages:
            raise ValueError(f"Page {end} does not exist. This PDF has "
                             f"{total_pages} page(s).")
        ranges.append((start, end))

    return ranges


def open_pdf(path):
    """Open a PDF safely. Handles PDFs that are locked only against editing."""
    reader = PdfReader(path)
    if reader.is_encrypted:
        try:
            ok = reader.decrypt("")
        except Exception:
            ok = False
        if not ok:
            raise ValueError(f"'{os.path.basename(path)}' is password protected "
                             "and cannot be opened.")
    return reader


def save_writer(writer, path):
    """Write a PdfWriter to disk."""
    with open(path, "wb") as handle:
        writer.write(handle)


def unique_path(path):
    """If the file already exists, add (2), (3)... so nothing is overwritten."""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    counter = 2
    while os.path.exists(f"{root} ({counter}){ext}"):
        counter += 1
    return f"{root} ({counter}){ext}"


def same_file(path_a, path_b):
    """True if two paths point to the same file."""
    norm = lambda p: os.path.normcase(os.path.abspath(p))
    return norm(path_a) == norm(path_b)


def build_writer(reader, entries):
    """
    entries = list of (page_index, extra_rotation).
    Rotation is applied to the copy in the new file, never to the original.
    """
    writer = PdfWriter()
    for index, rotation in entries:
        added = writer.add_page(reader.pages[index])
        if rotation:
            added.rotate(rotation)
    return writer


def describe_pages(numbers):
    """[1, 2, 3, 7, 5] -> '1-3, 7, 5'"""
    runs = []
    start = prev = numbers[0]
    for n in numbers[1:]:
        if n == prev + 1:
            prev = n
            continue
        runs.append((start, prev))
        start = prev = n
    runs.append((start, prev))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def file_label_for(numbers):
    """Name part for an output file: page_3, pages_3-7 or 5_pages."""
    if numbers == list(range(numbers[0], numbers[0] + len(numbers))):
        if len(numbers) == 1:
            return f"page_{numbers[0]}"
        return f"pages_{numbers[0]}-{numbers[-1]}"
    return f"{len(numbers)}_pages"


# ======================================================================
#  Page preview pictures (Organize & Split tool)
# ======================================================================

def render_preview(fitz_page, max_side=260):
    """Draw one PDF page as a small picture using PyMuPDF."""
    zoom = max_side / max(fitz_page.rect.width, fitz_page.rect.height)
    pix = fitz_page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    image = QImage(pix.samples, pix.width, pix.height, pix.stride,
                   QImage.Format_RGB888)
    return QPixmap.fromImage(image.copy())


def placeholder_preview(number):
    """Simple 'paper' picture with the page number, used without PyMuPDF."""
    pm = QPixmap(180, 240)
    pm.fill(QColor("#f6f7ff"))
    painter = QPainter(pm)
    painter.setPen(QPen(QColor("#dfe3f5"), 6))
    for y in range(40, 200, 22):
        painter.drawLine(24, y, 156, y)
    font = QFont("Segoe UI", 40)
    font.setBold(True)
    painter.setFont(font)
    painter.setPen(QColor("#9aa3c7"))
    painter.drawText(pm.rect(), Qt.AlignCenter, str(number))
    painter.end()
    return pm


def make_tile(preview, caption, color, rotation=0, split_after=False):
    """Compose the coloured card shown for one page in the grid."""
    if rotation:
        preview = preview.transformed(QTransform().rotate(rotation),
                                      Qt.SmoothTransformation)
    # Draw at the screen's real resolution so tiles stay sharp at 125%/150%
    app = QApplication.instance()
    dpr = max(1.0, app.devicePixelRatio()) if app else 1.0
    tile = QPixmap(int(TILE_W * dpr), int(TILE_H * dpr))
    tile.setDevicePixelRatio(dpr)
    tile.fill(Qt.transparent)
    p = QPainter(tile)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.SmoothPixmapTransform)

    # Card with coloured border (colour = which output file it belongs to)
    p.setPen(QPen(QColor(color), 4))
    p.setBrush(QColor("white"))
    p.drawRoundedRect(QRectF(3, 3, TILE_W - 6, TILE_H - 6), 12, 12)

    # Page picture
    area = QRect(14, 14, TILE_W - 28, TILE_H - 54)
    scaled = preview.scaled(area.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    x = area.x() + (area.width() - scaled.width()) // 2
    y = area.y() + (area.height() - scaled.height()) // 2
    p.drawPixmap(x, y, scaled)
    p.setPen(QPen(QColor("#d0d4e6"), 1))
    p.setBrush(Qt.NoBrush)
    p.drawRect(x, y, scaled.width() - 1, scaled.height() - 1)

    # Caption pill
    pill = QRect(14, TILE_H - 34, TILE_W - 28, 24)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(color))
    p.drawRoundedRect(QRectF(pill), 12, 12)
    font = QFont("Segoe UI", 9)
    font.setBold(True)
    p.setFont(font)
    p.setPen(QColor("white"))
    p.drawText(pill, Qt.AlignCenter, caption)

    # Split marker
    if split_after:
        badge = QRect(TILE_W - 74, 8, 66, 22)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#e53935"))
        p.drawRoundedRect(QRectF(badge), 8, 8)
        p.setPen(QColor("white"))
        font.setPointSize(8)
        p.setFont(font)
        p.drawText(badge, Qt.AlignCenter, "SPLIT  >")
        p.setPen(QPen(QColor("#e53935"), 5))
        p.drawLine(TILE_W - 4, 34, TILE_W - 4, TILE_H - 14)
    p.end()
    return tile


# ======================================================================
#  Small user interface helpers
# ======================================================================

def gradient_button(text, colors, big=False, compact=False):
    """A colourful rounded button."""
    c1, c2 = colors
    btn = QPushButton(text)
    btn.setCursor(Qt.PointingHandCursor)
    if big:
        padding, size = "14px 34px", 16
    elif compact:
        padding, size = "8px 6px", 12
    else:
        padding, size = "9px 18px", 13
    btn.setStyleSheet(f"""
        QPushButton {{
            color: white; border: none; border-radius: 12px;
            padding: {padding}; font-size: {size}px; font-weight: 700;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                        stop:0 {c1}, stop:1 {c2});
        }}
        QPushButton:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                        stop:0 {c2}, stop:1 {c1});
        }}
        QPushButton:disabled {{ background: #e3e5ee; color: #8a8fa8; }}
    """)
    # Keep the button at its natural height so text is never cut off
    btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
    btn.setMinimumHeight(btn.sizeHint().height())
    return btn


def header_banner(title, subtitle, colors):
    """Gradient banner shown at the top of each tool page."""
    c1, c2 = colors
    frame = QFrame()
    frame.setObjectName("Banner")
    frame.setStyleSheet(f"""
        #Banner {{
            border-radius: 18px;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                        stop:0 {c1}, stop:1 {c2});
        }}
        #Banner QLabel {{ color: white; background: transparent; }}
    """)
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(20, 8, 20, 8)
    layout.setSpacing(0)
    title_label = QLabel(title)
    title_label.setStyleSheet("font-size: 19px; font-weight: 800;")
    sub_label = QLabel(subtitle)
    sub_label.setStyleSheet("font-size: 12px;")
    sub_label.setWordWrap(True)
    layout.addWidget(title_label)
    layout.addWidget(sub_label)
    add_shadow(frame, QColor(c2))
    return frame


def white_card():
    """White rounded panel with a soft shadow."""
    card = QFrame()
    card.setObjectName("Card")
    add_shadow(card, QColor(60, 60, 120))
    return card


def add_shadow(widget, color):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(28)
    shadow.setOffset(0, 6)
    color.setAlpha(60)
    shadow.setColor(color)
    widget.setGraphicsEffect(shadow)


def show_error(parent, message):
    QMessageBox.warning(parent, APP_NAME, message)


def show_success(parent, message, folder):
    """Success dialog with a button that opens the output folder."""
    box = QMessageBox(parent)
    box.setWindowTitle(APP_NAME)
    box.setIcon(QMessageBox.Information)
    box.setText(message)
    open_btn = box.addButton("Open Folder", QMessageBox.ActionRole)
    box.addButton("Close", QMessageBox.RejectRole)
    box.exec_()
    if box.clickedButton() is open_btn:
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))


def set_status(widget, text):
    """Show a message in the main window's status bar."""
    window = widget.window()
    if isinstance(window, QMainWindow):
        window.statusBar().showMessage(text, 8000)


# ======================================================================
#  Merge page
# ======================================================================

class DropList(QListWidget):
    """List that accepts PDFs dragged from Explorer and can be reordered."""

    def __init__(self, on_files_dropped):
        super().__init__()
        self.on_files_dropped = on_files_dropped
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls()
                     if url.toLocalFile().lower().endswith(".pdf")]
            event.acceptProposedAction()
            self.on_files_dropped(paths)
        else:
            super().dropEvent(event)


class MergePage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)

        layout.addWidget(header_banner(
            "Merge PDFs",
            "Combine several PDF files into one. Drag the files in the list "
            "to set the order.",
            MERGE_COLORS))

        card = white_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 20, 24, 24)
        card_layout.setSpacing(14)

        hint = QLabel("Drag & drop PDF files here, or click 'Add PDFs'. "
                      "Drag items up or down to change the order.")
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        card_layout.addWidget(hint)

        self.file_list = DropList(self.add_paths)
        card_layout.addWidget(self.file_list, 1)

        # Row of small action buttons
        row = QHBoxLayout()
        row.setSpacing(10)
        add_btn = gradient_button("+ Add PDFs", ("#ff9966", "#ff5e62"))
        remove_btn = gradient_button("Remove", ("#f857a6", "#ff5858"))
        up_btn = gradient_button("Move Up", ("#36d1dc", "#5b86e5"))
        down_btn = gradient_button("Move Down", ("#5b86e5", "#36d1dc"))
        clear_btn = gradient_button("Clear All", ("#757f9a", "#5d6a8a"))
        add_btn.clicked.connect(self.browse_files)
        remove_btn.clicked.connect(self.remove_selected)
        up_btn.clicked.connect(lambda: self.move_selected(-1))
        down_btn.clicked.connect(lambda: self.move_selected(1))
        clear_btn.clicked.connect(self.clear_all)
        for btn in (add_btn, remove_btn, up_btn, down_btn, clear_btn):
            row.addWidget(btn)
        row.addStretch()
        card_layout.addLayout(row)

        self.count_label = QLabel("No files added yet.")
        self.count_label.setObjectName("Info")
        card_layout.addWidget(self.count_label)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        card_layout.addWidget(self.progress)

        self.merge_btn = gradient_button("Merge PDFs", MERGE_COLORS, big=True)
        self.merge_btn.clicked.connect(self.merge)
        card_layout.addWidget(self.merge_btn, 0, Qt.AlignRight)

        layout.addWidget(card, 1)
        self.update_count()

    # ---- list handling ----
    def browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF files", "", "PDF files (*.pdf)")
        self.add_paths(paths)

    def add_paths(self, paths):
        problems = []
        for path in paths:
            try:
                pages = len(open_pdf(path).pages)
            except Exception as err:
                problems.append(f"{os.path.basename(path)}: {err}")
                continue
            item = QListWidgetItem(f"{os.path.basename(path)}    "
                                   f"({pages} page{'s' if pages != 1 else ''})")
            item.setData(Qt.UserRole, path)
            item.setToolTip(path)
            self.file_list.addItem(item)
        self.update_count()
        if problems:
            show_error(self, "These files were skipped:\n\n" + "\n".join(problems))

    def remove_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self.update_count()

    def move_selected(self, step):
        row = self.file_list.currentRow()
        new_row = row + step
        if row < 0 or new_row < 0 or new_row >= self.file_list.count():
            return
        item = self.file_list.takeItem(row)
        self.file_list.insertItem(new_row, item)
        self.file_list.setCurrentRow(new_row)

    def clear_all(self):
        self.file_list.clear()
        self.update_count()

    def update_count(self):
        n = self.file_list.count()
        self.count_label.setText(
            "No files added yet." if n == 0 else f"{n} file(s) ready to merge.")

    # ---- merge ----
    def merge(self):
        paths = [self.file_list.item(i).data(Qt.UserRole)
                 for i in range(self.file_list.count())]
        if len(paths) < 2:
            show_error(self, "Please add at least two PDF files to merge.")
            return

        start_dir = os.path.join(os.path.dirname(paths[0]), "merged.pdf")
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save merged PDF as", start_dir, "PDF files (*.pdf)")
        if not out_path:
            return
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        if any(same_file(out_path, p) for p in paths):
            show_error(self, "Please choose a different name. The merged file "
                             "cannot replace one of the input files.")
            return

        self.progress.setVisible(True)
        self.progress.setRange(0, len(paths))
        self.merge_btn.setEnabled(False)
        try:
            writer = PdfWriter()
            for i, path in enumerate(paths, start=1):
                for page in open_pdf(path).pages:
                    writer.add_page(page)
                self.progress.setValue(i)
                QApplication.processEvents()
            save_writer(writer, out_path)
        except Exception as err:
            show_error(self, f"Merging failed:\n\n{err}")
            return
        finally:
            self.merge_btn.setEnabled(True)
            self.progress.setVisible(False)

        set_status(self, f"Merged {len(paths)} files into {os.path.basename(out_path)}")
        show_success(self, f"Done! {len(paths)} files were merged into:\n\n"
                           f"{out_path}", os.path.dirname(out_path))


# ======================================================================
#  Shared layout for Split and Extract pages
# ======================================================================

class RangeToolPage(QWidget):
    """Choose one PDF, type page ranges, press the action button."""

    def __init__(self, title, subtitle, colors, action_text, example_text):
        super().__init__()
        self.reader = None
        self.pdf_path = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)
        layout.addWidget(header_banner(title, subtitle, colors))

        card = white_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 24, 24, 24)
        card_layout.setSpacing(14)

        step1 = QLabel("Step 1 - Choose your PDF")
        step1.setObjectName("Step")
        card_layout.addWidget(step1)

        file_row = QHBoxLayout()
        choose_btn = gradient_button("Choose PDF", colors)
        choose_btn.clicked.connect(self.choose_file)
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("Info")
        file_row.addWidget(choose_btn)
        file_row.addSpacing(12)
        file_row.addWidget(self.file_label, 1)
        card_layout.addLayout(file_row)

        card_layout.addSpacing(8)
        step2 = QLabel("Step 2 - Enter the pages")
        step2.setObjectName("Step")
        card_layout.addWidget(step2)

        self.pages_edit = QLineEdit()
        self.pages_edit.setPlaceholderText("e.g.  1, 3-5, 8-10")
        self.pages_edit.setMinimumHeight(44)
        self.pages_edit.returnPressed.connect(self.run_action)
        card_layout.addWidget(self.pages_edit)

        example = QLabel(example_text)
        example.setObjectName("Hint")
        example.setWordWrap(True)
        card_layout.addWidget(example)

        card_layout.addStretch()
        self.action_btn = gradient_button(action_text, colors, big=True)
        self.action_btn.setEnabled(False)
        self.action_btn.clicked.connect(self.run_action)
        card_layout.addWidget(self.action_btn, 0, Qt.AlignRight)

        layout.addWidget(card, 1)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select a PDF", "", "PDF files (*.pdf)")
        if not path:
            return
        try:
            reader = open_pdf(path)
            total = len(reader.pages)
        except Exception as err:
            show_error(self, f"Could not open this PDF:\n\n{err}")
            return
        if total == 0:
            show_error(self, "This PDF has no pages.")
            return

        self.reader, self.pdf_path = reader, path
        self.file_label.setText(f"{os.path.basename(path)}   |   {total} page(s)")
        self.file_label.setToolTip(path)
        self.action_btn.setEnabled(True)
        self.pages_edit.setFocus()

    def get_ranges(self):
        """Return parsed ranges, or None after showing an error."""
        if self.reader is None:
            show_error(self, "Please choose a PDF first.")
            return None
        try:
            return parse_page_ranges(self.pages_edit.text(), len(self.reader.pages))
        except ValueError as err:
            show_error(self, str(err))
            return None

    def base_name(self):
        return os.path.splitext(os.path.basename(self.pdf_path))[0]

    def run_action(self):
        raise NotImplementedError


class SplitPage(RangeToolPage):
    def __init__(self):
        super().__init__(
            "Split PDF",
            "Break one PDF into several files - one file for each page or "
            "range you enter.",
            SPLIT_COLORS, "Split PDF",
            "Example: for a 10-page PDF, entering  1, 3-5, 8-10  creates 3 files: "
            "page 1 | pages 3-5 | pages 8-10.")

    def run_action(self):
        ranges = self.get_ranges()
        if not ranges:
            return
        folder = QFileDialog.getExistingDirectory(
            self, "Choose a folder for the split files",
            os.path.dirname(self.pdf_path))
        if not folder:
            return

        created = []
        try:
            for number, (start, end) in enumerate(ranges, start=1):
                writer = PdfWriter()
                for index in range(start - 1, end):
                    writer.add_page(self.reader.pages[index])
                label = f"page_{start}" if start == end else f"pages_{start}-{end}"
                out_path = unique_path(os.path.join(
                    folder, f"{self.base_name()}_part{number}_{label}.pdf"))
                save_writer(writer, out_path)
                created.append(os.path.basename(out_path))
        except Exception as err:
            show_error(self, f"Splitting failed:\n\n{err}")
            return

        set_status(self, f"Created {len(created)} file(s) in {folder}")
        show_success(self, f"Done! {len(created)} file(s) created:\n\n"
                           + "\n".join(created), folder)


class ExtractPage(RangeToolPage):
    def __init__(self):
        super().__init__(
            "Extract Pages",
            "Pick the pages you need and save them together as one new PDF.",
            EXTRACT_COLORS, "Extract Pages",
            "Example: for a 10-page PDF, entering  1, 3-5, 8-10  creates one file "
            "with pages 1, 3, 4, 5, 8, 9, 10 (in the order you type them).")

    def run_action(self):
        ranges = self.get_ranges()
        if not ranges:
            return
        start_path = os.path.join(os.path.dirname(self.pdf_path),
                                  f"{self.base_name()}_extracted.pdf")
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save extracted pages as", start_path, "PDF files (*.pdf)")
        if not out_path:
            return
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        if same_file(out_path, self.pdf_path):
            show_error(self, "Please choose a different name. The new file "
                             "cannot replace the original PDF.")
            return

        page_numbers = []
        for start, end in ranges:
            page_numbers.extend(range(start, end + 1))

        try:
            writer = PdfWriter()
            for number in page_numbers:
                writer.add_page(self.reader.pages[number - 1])
            save_writer(writer, out_path)
        except Exception as err:
            show_error(self, f"Extraction failed:\n\n{err}")
            return

        pages_text = ", ".join(str(n) for n in page_numbers)
        set_status(self, f"Extracted {len(page_numbers)} page(s)")
        show_success(self, f"Done! Pages {pages_text} were saved to:\n\n{out_path}",
                     os.path.dirname(out_path))


# ======================================================================
#  Organize & Split page
# ======================================================================

class PageGrid(QListWidget):
    """Grid of page tiles. Pages can be dragged to a new position."""

    orderChanged = pyqtSignal()
    filesDropped = pyqtSignal(list)
    deletePressed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setObjectName("PageGrid")
        self.setViewMode(QListView.ListMode)
        self.setFlow(QListView.LeftToRight)
        self.setWrapping(True)
        self.setResizeMode(QListView.Adjust)
        self.setUniformItemSizes(True)
        self.setIconSize(QSize(TILE_W, TILE_H))
        self.setGridSize(QSize(TILE_W + 14, TILE_H + 14))
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setDropIndicatorShown(True)
        self.setMinimumSize(TILE_W + 60, TILE_H + 60)

    # ---- our own drag & drop so pages are only ever moved, never lost ----
    def startDrag(self, actions):
        items = self.selectedItems()
        if not items:
            return
        drag = QDrag(self)
        drag.setMimeData(self.mimeData(items))
        drag.setPixmap(items[0].icon().pixmap(QSize(TILE_W // 2, TILE_H // 2)))
        self.setState(QAbstractItemView.DraggingState)
        drag.exec_(Qt.MoveAction)
        self.setState(QAbstractItemView.NoState)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.source() is self:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        elif event.source() is self:
            super().dragMoveEvent(event)
            event.setDropAction(Qt.MoveAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [u.toLocalFile() for u in event.mimeData().urls()
                     if u.toLocalFile().lower().endswith(".pdf")]
            event.acceptProposedAction()
            if paths:
                self.filesDropped.emit(paths)
            return
        if event.source() is not self:
            event.ignore()
            return
        rows = [self.row(item) for item in self.selectedItems()]
        target = self.target_row(event.pos())
        event.setDropAction(Qt.IgnoreAction)  # we move the items ourselves
        event.accept()
        self.move_rows(rows, target)

    def target_row(self, pos):
        """Insert position nearest to the mouse pointer."""
        best_row, best_dist, best_rect = self.count(), None, None
        for row in range(self.count()):
            rect = self.visualItemRect(self.item(row))
            dist = (rect.center() - pos).manhattanLength()
            if best_dist is None or dist < best_dist:
                best_row, best_dist, best_rect = row, dist, rect
        if best_rect is None:
            return 0
        return best_row + (1 if pos.x() > best_rect.center().x() else 0)

    def move_rows(self, rows, target):
        """Move the given rows (kept in order) so they start at 'target'."""
        rows = sorted(set(rows))
        if not rows:
            return
        target -= sum(1 for r in rows if r < target)
        taken = [self.takeItem(r) for r in reversed(rows)][::-1]
        self.clearSelection()
        target = max(0, min(target, self.count()))
        for offset, item in enumerate(taken):
            self.insertItem(target + offset, item)
            item.setSelected(True)
        self.scrollToItem(taken[0])
        self.orderChanged.emit()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            self.deletePressed.emit()
        else:
            super().keyPressEvent(event)


class OrganizePage(QWidget):
    def __init__(self):
        super().__init__()
        self.reader = None
        self.pdf_path = None
        self.previews = []
        self.tool_buttons = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)
        layout.addWidget(header_banner(
            "Organize & Split",
            "Drag pages to reorder, rotate, delete or sort. "
            "Double-click a page to split after it.",
            ORGANIZE_COLORS))

        self.card = white_card()
        outer = QHBoxLayout(self.card)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(12)

        # ---------------- left: file + page grid ----------------
        left = QVBoxLayout()
        left.setSpacing(10)
        file_row = QHBoxLayout()
        choose_btn = gradient_button("Choose PDF", ORGANIZE_COLORS)
        choose_btn.clicked.connect(self.choose_file)
        self.file_label = QLabel("No file selected - choose a PDF or drop one below")
        self.file_label.setObjectName("Info")
        file_row.addWidget(choose_btn)
        file_row.addSpacing(10)
        file_row.addWidget(self.file_label, 1)
        left.addLayout(file_row)

        if not HAS_PREVIEWS:
            tip = QLabel("Tip: run  pip install pymupdf  to see real page "
                         "previews. Numbered placeholders are shown for now.")
            tip.setObjectName("Hint")
            tip.setWordWrap(True)
            left.addWidget(tip)

        self.grid = PageGrid()
        self.grid.orderChanged.connect(self.refresh_tiles)
        self.grid.filesDropped.connect(lambda paths: self.load_pdf(paths[0]))
        self.grid.deletePressed.connect(self.delete_selected)
        self.grid.itemDoubleClicked.connect(self.toggle_one)
        left.addWidget(self.grid, 1)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        left.addWidget(self.progress)

        self.summary = QLabel("")
        self.summary.setObjectName("Info")
        self.summary.setWordWrap(True)
        left.addWidget(self.summary)

        self.save_btn = gradient_button("💾 Save as One PDF", ("#43a047", "#2e7d32"), compact=True)
        self.split_btn = gradient_button("✂ Split at Marks", ("#e53935", "#c2185b"), compact=True)
        self.save_btn.clicked.connect(self.save_one)
        self.split_btn.clicked.connect(self.split_at_marks)
        outer.addLayout(left, 1)

        # ---------------- right: tool panel ----------------
        panel = QFrame()
        panel.setObjectName("ToolPanel")
        panel.setFixedWidth(250)
        side = QVBoxLayout(panel)
        side.setContentsMargins(12, 10, 12, 10)
        side.setSpacing(5)

        blue, purple = ("#36d1dc", "#5b86e5"), ("#8e54e9", "#4776e6")
        green, grey = ("#11998e", "#38ef7d"), ("#757f9a", "#5d6a8a")
        red, orange = ("#f857a6", "#ff5858"), ("#ff9a00", "#ff5e00")

        side.addWidget(self.section("Arrange selected pages"))
        side.addLayout(self.button_grid([
            ("< Move Left", blue, lambda: self.move_selected(-1)),
            ("Move Right >", blue, lambda: self.move_selected(1)),
            ("Rotate Left", purple, lambda: self.rotate_selected(-90)),
            ("Rotate Right", purple, lambda: self.rotate_selected(90)),
            ("Delete", red, self.delete_selected),
            ("Select All", grey, self.grid.selectAll),
        ]))

        side.addWidget(self.section("Sort all pages"))
        side.addLayout(self.button_grid([
            ("Reverse", green, self.reverse_order),
            ("Sort 1 to N", green, self.sort_ascending),
            ("Odd, then Even", green, self.odd_then_even),
            ("Reset All", grey, self.reset_pages),
        ]))

        side.addWidget(self.section("Split marks"))
        side.addLayout(self.button_grid([
            ("Toggle Split", orange, self.toggle_selected),
            ("Clear Marks", grey, self.clear_marks),
        ]))
        every_row = QHBoxLayout()
        every_label = QLabel("Every")
        every_label.setObjectName("Hint")
        self.every_spin = QSpinBox()
        self.every_spin.setRange(1, 9999)
        self.every_spin.setValue(2)
        every_btn = gradient_button("pages", orange, compact=True)
        every_btn.setToolTip("Put a split mark after every N pages")
        every_btn.clicked.connect(self.split_every_n)
        self.tool_buttons.append(every_btn)
        every_row.addWidget(every_label)
        every_row.addWidget(self.every_spin, 1)
        every_row.addWidget(every_btn, 1)
        side.addLayout(every_row)

        side.addWidget(self.section("Save"))
        side.addWidget(self.save_btn)
        side.addWidget(self.split_btn)
        side.addStretch()
        panel.setToolTip("Drag pages to move them. Ctrl/Shift + click selects several. "
                         "Press Delete to remove. Each border colour = one output file.")
        outer.addWidget(panel)

        layout.addWidget(self.card, 1)
        self.set_tools_enabled(False)
        self.update_summary()

    # ---------------- small layout helpers ----------------
    @staticmethod
    def section(text):
        label = QLabel(text)
        label.setObjectName("Step")
        return label

    def button_grid(self, specs, columns=2):
        grid = QGridLayout()
        grid.setSpacing(6)
        for i, (text, colors, handler) in enumerate(specs):
            btn = gradient_button(text, colors, compact=True)
            btn.clicked.connect(handler)
            self.tool_buttons.append(btn)
            grid.addWidget(btn, i // columns, i % columns)
        return grid

    def set_tools_enabled(self, enabled):
        for btn in self.tool_buttons + [self.save_btn, self.split_btn]:
            btn.setEnabled(enabled)
        self.every_spin.setEnabled(enabled)

    # ---------------- loading ----------------
    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select a PDF", "", "PDF files (*.pdf)")
        if path:
            self.load_pdf(path)

    def load_pdf(self, path):
        try:
            reader = open_pdf(path)
            total = len(reader.pages)
        except Exception as err:
            show_error(self, f"Could not open this PDF:\n\n{err}")
            return
        if total == 0:
            show_error(self, "This PDF has no pages.")
            return

        self.card.setEnabled(False)
        self.progress.setRange(0, total)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            previews = self.make_previews(path, total)
        finally:
            QApplication.restoreOverrideCursor()
            self.progress.setVisible(False)
            self.card.setEnabled(True)

        self.reader, self.pdf_path, self.previews = reader, path, previews
        self.file_label.setText(f"{os.path.basename(path)}   |   {total} page(s)")
        self.file_label.setToolTip(path)
        self.reset_pages()
        self.set_tools_enabled(True)
        set_status(self, f"Loaded {os.path.basename(path)}")

    def make_previews(self, path, total):
        previews = []
        if HAS_PREVIEWS:
            try:
                doc = fitz.open(path)
                if doc.needs_pass:
                    doc.authenticate("")
                for i in range(total):
                    previews.append(render_preview(doc[i]))
                    if i % 4 == 0:
                        self.progress.setValue(i)
                        QApplication.processEvents()
                doc.close()
            except Exception:
                previews = []
        if len(previews) != total:
            previews = [placeholder_preview(i + 1) for i in range(total)]
        return previews

    # ---------------- page list state ----------------
    def states(self):
        """List of (page_index, rotation, split_after) in current order."""
        result = []
        for row in range(self.grid.count()):
            item = self.grid.item(row)
            result.append((item.data(ROLE_PAGE), item.data(ROLE_ROT),
                           bool(item.data(ROLE_BREAK))))
        return result

    def set_states(self, states):
        self.grid.clear()
        for page, rotation, split_after in states:
            item = QListWidgetItem()
            item.setData(ROLE_PAGE, page)
            item.setData(ROLE_ROT, rotation)
            item.setData(ROLE_BREAK, split_after)
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled)
            self.grid.addItem(item)
        self.refresh_tiles()

    def refresh_tiles(self):
        """Redraw every tile: colours show which output file each page goes to."""
        count = self.grid.count()
        self.grid.setUpdatesEnabled(False)
        segment = 0
        for row in range(count):
            item = self.grid.item(row)
            page = item.data(ROLE_PAGE)
            rotation = item.data(ROLE_ROT) or 0
            split_after = bool(item.data(ROLE_BREAK)) and row < count - 1
            color = SEGMENT_COLORS[segment % len(SEGMENT_COLORS)]
            caption = f"Page {page + 1}" + (f"  ({rotation}\u00b0)" if rotation else "")
            tile = make_tile(self.previews[page], caption, color, rotation, split_after)
            icon = QIcon()
            icon.addPixmap(tile, QIcon.Normal)
            icon.addPixmap(tile, QIcon.Selected)
            item.setIcon(icon)
            item.setToolTip(f"Position {row + 1}  |  original page {page + 1}\n"
                            "Double-click to add/remove a split after this page")
            if split_after:
                segment += 1
        self.grid.setUpdatesEnabled(True)
        self.update_summary()

    def segments(self):
        """Split the current order at the marks -> list of [(page, rotation), ...]."""
        groups, current = [], []
        states = self.states()
        for row, (page, rotation, split_after) in enumerate(states):
            current.append((page, rotation))
            if split_after and row < len(states) - 1:
                groups.append(current)
                current = []
        if current:
            groups.append(current)
        return groups

    def update_summary(self):
        count = self.grid.count()
        if count == 0:
            self.summary.setText("")
            return
        groups = self.segments()
        if len(groups) == 1:
            self.summary.setText(f"{count} page(s). No split marks yet - "
                                 "double-click a page to split after it.")
            return
        parts = [f"[{describe_pages([p + 1 for p, _ in g])}]" for g in groups]
        shown = "  ".join(parts[:8]) + ("  ..." if len(parts) > 8 else "")
        self.summary.setText(f"{len(groups)} files will be created:  {shown}")

    def selected_rows(self):
        rows = sorted(self.grid.row(i) for i in self.grid.selectedItems())
        if not rows:
            set_status(self, "Select one or more pages first.")
        return rows

    # ---------------- arrange ----------------
    def move_selected(self, step):
        rows = self.selected_rows()
        if not rows:
            return
        if step < 0 and rows[0] > 0:
            self.grid.move_rows(rows, rows[0] - 1)
        elif step > 0 and rows[-1] < self.grid.count() - 1:
            self.grid.move_rows(rows, rows[-1] + 2)

    def rotate_selected(self, degrees):
        rows = self.selected_rows()
        for row in rows:
            item = self.grid.item(row)
            item.setData(ROLE_ROT, ((item.data(ROLE_ROT) or 0) + degrees) % 360)
        if rows:
            self.refresh_tiles()

    def delete_selected(self):
        rows = self.selected_rows()
        if not rows:
            return
        if len(rows) == self.grid.count():
            show_error(self, "At least one page must remain.")
            return
        for row in reversed(rows):
            self.grid.takeItem(row)
        self.refresh_tiles()
        set_status(self, f"Removed {len(rows)} page(s) from the list "
                         "(the original file is not changed).")

    # ---------------- sort ----------------
    def reverse_order(self):
        self.set_states(self.states()[::-1])

    def sort_ascending(self):
        self.set_states(sorted(self.states(), key=lambda s: s[0]))

    def odd_then_even(self):
        states = self.states()
        self.set_states(states[0::2] + states[1::2])

    def reset_pages(self):
        self.set_states([(i, 0, False) for i in range(len(self.previews))])

    # ---------------- split marks ----------------
    def toggle_one(self, item):
        if self.grid.row(item) == self.grid.count() - 1:
            set_status(self, "A split after the last page has no effect.")
            return
        item.setData(ROLE_BREAK, not item.data(ROLE_BREAK))
        self.refresh_tiles()

    def toggle_selected(self):
        rows = self.selected_rows()
        for row in rows:
            item = self.grid.item(row)
            item.setData(ROLE_BREAK, not item.data(ROLE_BREAK))
        if rows:
            self.refresh_tiles()

    def split_every_n(self):
        n = self.every_spin.value()
        for row in range(self.grid.count()):
            self.grid.item(row).setData(ROLE_BREAK, (row + 1) % n == 0)
        self.refresh_tiles()

    def clear_marks(self):
        for row in range(self.grid.count()):
            self.grid.item(row).setData(ROLE_BREAK, False)
        self.refresh_tiles()

    # ---------------- output ----------------
    def base_name(self):
        return os.path.splitext(os.path.basename(self.pdf_path))[0]

    def save_one(self):
        if self.reader is None or self.grid.count() == 0:
            show_error(self, "Please choose a PDF first.")
            return
        start_path = os.path.join(os.path.dirname(self.pdf_path),
                                  f"{self.base_name()}_organized.pdf")
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save organized PDF as", start_path, "PDF files (*.pdf)")
        if not out_path:
            return
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        if same_file(out_path, self.pdf_path):
            show_error(self, "Please choose a different name. The new file "
                             "cannot replace the original PDF.")
            return
        entries = [(page, rot) for page, rot, _ in self.states()]
        try:
            save_writer(build_writer(self.reader, entries), out_path)
        except Exception as err:
            show_error(self, f"Saving failed:\n\n{err}")
            return
        set_status(self, f"Saved {len(entries)} page(s) to {os.path.basename(out_path)}")
        show_success(self, f"Done! {len(entries)} page(s) saved in your new "
                           f"order to:\n\n{out_path}", os.path.dirname(out_path))

    def split_at_marks(self):
        if self.reader is None or self.grid.count() == 0:
            show_error(self, "Please choose a PDF first.")
            return
        groups = self.segments()
        if len(groups) < 2:
            show_error(self, "There are no split marks yet.\n\nDouble-click a page "
                             "(or select it and press 'Toggle Split After Selected') "
                             "to split the PDF after that page.")
            return
        folder = QFileDialog.getExistingDirectory(
            self, "Choose a folder for the split files", os.path.dirname(self.pdf_path))
        if not folder:
            return
        created = []
        try:
            for number, group in enumerate(groups, start=1):
                label = file_label_for([page + 1 for page, _ in group])
                out_path = unique_path(os.path.join(
                    folder, f"{self.base_name()}_part{number}_{label}.pdf"))
                save_writer(build_writer(self.reader, group), out_path)
                created.append(os.path.basename(out_path))
        except Exception as err:
            show_error(self, f"Splitting failed:\n\n{err}")
            return
        set_status(self, f"Created {len(created)} file(s) in {folder}")
        listing = "\n".join(created[:15]) + ("\n..." if len(created) > 15 else "")
        show_success(self, f"Done! {len(created)} file(s) created:\n\n{listing}", folder)


# ======================================================================
#  OCR (read text from scanned PDFs and images)
# ======================================================================

_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()


def get_ocr_engine():
    """One shared OCR engine for every tool (it is slow to start)."""
    global _OCR_ENGINE
    with _OCR_LOCK:
        if _OCR_ENGINE is None:
            from rapidocr_onnxruntime import RapidOCR
            _OCR_ENGINE = RapidOCR()
        return _OCR_ENGINE


def open_document(path):
    """Open a PDF or an image as a PDF document (PyMuPDF)."""
    doc = fitz.open(path)
    if doc.needs_pass and not doc.authenticate(""):
        raise ValueError(f"'{os.path.basename(path)}' is password protected.")
    if not doc.is_pdf:
        doc = fitz.open("pdf", doc.convert_to_pdf())
    return doc


def read_page_text(page, allow_ocr=True):
    """Text of a page. Scanned pages are read with OCR when available.
    Returns (text, used_ocr)."""
    text = page.get_text("text", sort=True).strip()
    if len(text) >= 20 or not (allow_ocr and HAS_OCR):
        return text, False
    return boxes_to_text(ocr_page(get_ocr_engine(), page)), True


class TaskWorker(QThread):
    """Runs func(worker, *args) in the background.
    func calls worker.report(done, total, message) and checks worker.cancelled."""

    progress = pyqtSignal(int, int, str)
    done = pyqtSignal(object)
    failed = pyqtSignal(str)
    running = set()

    def __init__(self, func, *args):
        super().__init__()
        self.func, self.args = func, args
        self.cancelled = False
        TaskWorker.running.add(self)
        self.finished.connect(lambda: TaskWorker.running.discard(self))

    def report(self, done, total, message=""):
        self.progress.emit(done, total, message)

    def run(self):
        try:
            result = self.func(self, *self.args)
        except Exception as err:
            self.failed.emit(str(err) or err.__class__.__name__)
            return
        self.done.emit(result)


def ocr_page(engine, page):
    """Read one page. Returns [(fitz.Rect in page units, text), ...]."""
    zoom = min(2.5, 3200 / max(page.rect.width, page.rect.height))
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    import numpy as np  # installed together with the OCR engine
    image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)
    result, _ = engine(image[:, :, ::-1].copy())  # RGB -> BGR
    boxes = []
    for box, text, _score in result or []:
        xs = [pt[0] / zoom for pt in box]
        ys = [pt[1] / zoom for pt in box]
        boxes.append((fitz.Rect(min(xs), min(ys), max(xs), max(ys)), text))
    return boxes


def boxes_to_text(boxes):
    """Arrange OCR boxes into reading order: top to bottom, left to right."""
    lines = []
    for rect, text in sorted(boxes, key=lambda b: (b[0].y0 + b[0].y1) / 2):
        centre = (rect.y0 + rect.y1) / 2
        if lines and abs(centre - lines[-1]["centre"]) < rect.height * 0.6:
            lines[-1]["items"].append((rect.x0, text))
        else:
            lines.append({"centre": centre, "items": [(rect.x0, text)]})
    return "\n".join("   ".join(t for _, t in sorted(line["items"]))
                     for line in lines)


def add_text_layer(page, boxes):
    """Put invisible, searchable text on top of the scanned image."""
    for rect, text in boxes:
        try:
            size = rect.height * 0.8
            width_at_1 = fitz.get_text_length(text, fontname="helv", fontsize=1)
            if width_at_1 > 0:
                size = min(size, rect.width / width_at_1)
            if size < 1:
                continue
            point = fitz.Point(rect.x0, rect.y1 - rect.height * 0.2) * page.derotation_matrix
            page.insert_text(point, text, fontsize=size, fontname="helv",
                             render_mode=3, rotate=page.rotation)
        except Exception:
            pass  # one odd word must never stop the whole document


class OcrWorker(QThread):
    """Runs OCR in the background so the window stays responsive."""

    status = pyqtSignal(str)
    progress = pyqtSignal(int, int)
    page_done = pyqtSignal(int, str)
    finished_ok = pyqtSignal(bool)      # True if cancelled
    failed = pyqtSignal(str)

    def __init__(self, path, pages, skip_text_pages, pdf_out=None):
        super().__init__()
        self.path = path
        self.pages = pages
        self.skip_text_pages = skip_text_pages
        self.pdf_out = pdf_out
        self.cancelled = False

    def run(self):
        try:
            self.status.emit("Starting the OCR engine (first time takes a few seconds)...")
            engine = get_ocr_engine()

            src = fitz.open(self.path)
            if src.needs_pass and not src.authenticate(""):
                raise ValueError("This file is password protected.")
            if not src.is_pdf:                      # image -> PDF page
                src = fitz.open("pdf", src.convert_to_pdf())
            out = fitz.open() if self.pdf_out else None

            total = len(self.pages)
            for done, index in enumerate(self.pages, start=1):
                if self.cancelled:
                    break
                self.status.emit(f"Reading page {index + 1}  ({done} of {total})...")
                page = src[index]
                existing = page.get_text().strip()
                if self.skip_text_pages and existing:
                    text, boxes = existing, []
                else:
                    boxes = ocr_page(engine, page)
                    text = boxes_to_text(boxes)
                if out is not None:
                    out.insert_pdf(src, from_page=index, to_page=index)
                    add_text_layer(out[-1], boxes)
                self.page_done.emit(index + 1, text)
                self.progress.emit(done, total)

            if out is not None and not self.cancelled:
                out.save(self.pdf_out, garbage=3, deflate=True)
            self.finished_ok.emit(self.cancelled)
        except Exception as err:
            self.failed.emit(str(err))


class OcrPage(QWidget):
    def __init__(self):
        super().__init__()
        self.path = None
        self.total = 0
        self.worker = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)
        layout.addWidget(header_banner(
            "OCR - Read Text",
            "Turn scanned PDFs or photos into text you can copy, search and save. "
            "Works fully offline.",
            OCR_COLORS))

        card = white_card()
        col = QVBoxLayout(card)
        col.setContentsMargins(24, 20, 24, 20)
        col.setSpacing(10)

        if not HAS_OCR:
            warn = QLabel("OCR needs two extra libraries. Close the app, run this in "
                          "Command Prompt, then start it again:\n\n"
                          "    pip install pymupdf rapidocr_onnxruntime")
            warn.setObjectName("Info")
            warn.setWordWrap(True)
            warn.setTextInteractionFlags(Qt.TextSelectableByMouse)
            col.addWidget(warn)

        col.addWidget(self.step("Step 1 - Choose a scanned PDF or an image"))
        row = QHBoxLayout()
        self.choose_btn = gradient_button("Choose File", OCR_COLORS)
        self.choose_btn.clicked.connect(self.choose_file)
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("Info")
        row.addWidget(self.choose_btn)
        row.addSpacing(10)
        row.addWidget(self.file_label, 1)
        col.addLayout(row)

        col.addWidget(self.step("Step 2 - Pages to read (optional), then run"))
        self.pages_edit = QLineEdit()
        self.pages_edit.setPlaceholderText("Leave empty for all pages, or e.g.  1-3, 5")
        self.pages_edit.setMinimumHeight(40)
        self.skip_check = QCheckBox("Skip pages that already have text")
        self.skip_check.setChecked(True)
        run_row = QHBoxLayout()
        run_row.addWidget(self.pages_edit, 1)
        run_row.addWidget(self.skip_check)
        self.read_btn = gradient_button("Read Text", OCR_COLORS)
        self.pdf_btn = gradient_button("Make Searchable PDF", ("#4776e6", "#8e54e9"))
        self.cancel_btn = gradient_button("Cancel", ("#757f9a", "#5d6a8a"))
        self.read_btn.clicked.connect(lambda: self.start(make_pdf=False))
        self.pdf_btn.clicked.connect(lambda: self.start(make_pdf=True))
        self.cancel_btn.clicked.connect(self.cancel)
        for btn in (self.read_btn, self.pdf_btn, self.cancel_btn):
            run_row.addWidget(btn)
        col.addLayout(run_row)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        col.addWidget(self.progress)
        self.status_label = QLabel("")
        self.status_label.setObjectName("Hint")
        col.addWidget(self.status_label)

        self.text_box = QPlainTextEdit()
        self.text_box.setPlaceholderText("The text found in your file will appear here.")
        self.text_box.setMinimumHeight(140)
        col.addWidget(self.text_box, 1)

        out_row = QHBoxLayout()
        out_row.addStretch()
        self.copy_btn = gradient_button("Copy Text", ("#36d1dc", "#5b86e5"))
        self.save_txt_btn = gradient_button("Save as .txt", ("#11998e", "#38ef7d"))
        self.copy_btn.clicked.connect(self.copy_text)
        self.save_txt_btn.clicked.connect(self.save_text)
        out_row.addWidget(self.copy_btn)
        out_row.addWidget(self.save_txt_btn)
        col.addLayout(out_row)

        layout.addWidget(card, 1)
        self.set_running(False)

    @staticmethod
    def step(text):
        label = QLabel(text)
        label.setObjectName("Step")
        return label

    def set_running(self, running):
        ready = HAS_OCR and self.path is not None and not running
        self.read_btn.setEnabled(ready)
        self.pdf_btn.setEnabled(ready)
        self.choose_btn.setEnabled(HAS_OCR and not running)
        self.cancel_btn.setEnabled(running)
        self.cancel_btn.setVisible(running)
        self.progress.setVisible(running)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select a scanned PDF or image", "",
            "PDF or image (*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp)")
        if not path:
            return
        try:
            doc = fitz.open(path)
            if doc.needs_pass and not doc.authenticate(""):
                raise ValueError("This file is password protected.")
            total = len(doc)
            doc.close()
        except Exception as err:
            show_error(self, f"Could not open this file:\n\n{err}")
            return
        if total == 0:
            show_error(self, "This file has no pages.")
            return
        self.path, self.total = path, total
        self.file_label.setText(f"{os.path.basename(path)}   |   {total} page(s)")
        self.file_label.setToolTip(path)
        self.set_running(False)

    def start(self, make_pdf):
        if self.path is None:
            show_error(self, "Please choose a file first.")
            return
        text = self.pages_edit.text().strip()
        if text:
            try:
                ranges = parse_page_ranges(text, self.total)
            except ValueError as err:
                show_error(self, str(err))
                return
            pages = []
            for start, end in ranges:
                pages.extend(i - 1 for i in range(start, end + 1))
        else:
            pages = list(range(self.total))

        pdf_out = None
        if make_pdf:
            base = os.path.splitext(os.path.basename(self.path))[0]
            start_path = os.path.join(os.path.dirname(self.path), f"{base}_searchable.pdf")
            pdf_out, _ = QFileDialog.getSaveFileName(
                self, "Save searchable PDF as", start_path, "PDF files (*.pdf)")
            if not pdf_out:
                return
            if not pdf_out.lower().endswith(".pdf"):
                pdf_out += ".pdf"
            if same_file(pdf_out, self.path):
                show_error(self, "Please choose a different name. The new file "
                                 "cannot replace the original.")
                return

        self.text_box.clear()
        self.progress.setRange(0, len(pages))
        self.progress.setValue(0)
        self.worker = OcrWorker(self.path, pages, self.skip_check.isChecked(), pdf_out)
        self.worker.status.connect(self.status_label.setText)
        self.worker.progress.connect(lambda done, _total: self.progress.setValue(done))
        self.worker.page_done.connect(self.add_page_text)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.set_running(True)
        self.worker.start()

    def add_page_text(self, page_number, text):
        self.text_box.appendPlainText(f"===== Page {page_number} =====")
        self.text_box.appendPlainText(text if text else "(no text found)")
        self.text_box.appendPlainText("")

    def on_finished(self, cancelled):
        pdf_out = self.worker.pdf_out
        self.set_running(False)
        if cancelled:
            self.status_label.setText("Cancelled. Text read so far is shown below.")
            return
        self.status_label.setText("Finished.")
        set_status(self, "OCR finished")
        if pdf_out:
            show_success(self, "Done! Your searchable PDF was saved to:\n\n"
                               f"{pdf_out}\n\nYou can now search and copy its text "
                               "in any PDF reader.", os.path.dirname(pdf_out))

    def on_failed(self, message):
        self.set_running(False)
        self.status_label.setText("OCR stopped because of an error.")
        show_error(self, f"OCR failed:\n\n{message}")

    def cancel(self):
        if self.worker is not None:
            self.worker.cancelled = True
            self.status_label.setText("Cancelling after the current page...")

    def shutdown(self):
        """Called when the window closes."""
        if self.worker is not None and self.worker.isRunning():
            self.worker.cancelled = True
            self.worker.wait(15000)

    def copy_text(self):
        QApplication.clipboard().setText(self.text_box.toPlainText())
        set_status(self, "Text copied to clipboard.")

    def save_text(self):
        content = self.text_box.toPlainText()
        if not content.strip():
            show_error(self, "There is no text to save yet.")
            return
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "ocr"
        start_dir = os.path.dirname(self.path) if self.path else ""
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save text as", os.path.join(start_dir, f"{base}_text.txt"),
            "Text files (*.txt)")
        if not out_path:
            return
        try:
            with open(out_path, "w", encoding="utf-8") as handle:
                handle.write(content)
        except Exception as err:
            show_error(self, f"Saving failed:\n\n{err}")
            return
        show_success(self, f"Text saved to:\n\n{out_path}", os.path.dirname(out_path))


# ======================================================================
#  Shared patterns (summary, invoices, redaction)
# ======================================================================

MONTHS = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?"
DATE_PATTERN = (r"(?:\d{1,2}[/.-]\d{1,2}[/.-](?:19|20)?\d{2}"
                r"|\d{1,2}(?:st|nd|rd|th)?[ -]" + MONTHS + r",?[ -](?:19|20)?\d{2}"
                r"|" + MONTHS + r"[ ]\d{1,2},?[ ](?:19|20)\d{2}"
                r"|(?:19|20)\d{2}-\d{2}-\d{2})")
DATE_RE = re.compile(r"(?<![\d/.-])" + DATE_PATTERN + r"(?![\d/])")
MONEY_RE = re.compile(r"(?:₹|Rs\.?|INR|\$|USD|£|GBP|€|EUR)[ ]?\d[\d,]*(?:\.\d{1,2})?"
                      r"(?:[ ]?(?:lakhs?|lacs?|crores?|cr|million|mn|k)\b)?", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
GSTIN_RE = re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b")
COMPANY_RE = re.compile(
    r"\b(?:[A-Z][A-Za-z0-9&'.-]*[ ]){1,6}(?:Private[ ]Limited|Pvt\.?[ ]?Ltd\.?|Limited|"
    r"Ltd\.?|LLP|LLC|Inc\.?|Corp(?:oration)?\.?|Enterprises|Industries)(?![A-Za-z])")
REF_RE = re.compile(
    r"\b(Policy|Invoice|Order|Reference|Ref|Account|Customer|Proposal|Certificate|Claim|"
    r"PO|Contract|Agreement|Member|Employee)\s*(?:No\.?|Number|#|ID)\s*[:\-]?\s*"
    r"([A-Z0-9][A-Z0-9/\-]{3,})", re.I)


def money_value(text):
    try:
        value = float(re.sub(r"[^\d.]", "", re.sub(r"(?i)(rs\.|inr|usd|gbp|eur)", "", text)))
    except ValueError:
        return 0.0
    if re.search(r"(?i)crore|\bcr\b", text):
        value *= 1e7
    elif re.search(r"(?i)lakh|lac", text):
        value *= 1e5
    return value


# ======================================================================
#  1-page Summary
# ======================================================================

STOPWORDS = set("""a about above after again against all am an and any are as at be because been
before being below between both but by can could did do does doing down during each few for from
further had has have having he her here hers herself him himself his how i if in into is it its
itself just me more most my myself no nor not now of off on once only or other our ours ourselves
out over own same she should so some such than that the their theirs them themselves then there
these they this those through to too under until up very was we were what when where which while
who whom why will with you your yours yourself also may shall would upon within without per via etc
page any such shall said thereof herein hereby whereas one two three must""".split())

CUE_WORDS = re.compile(r"(?i)\b(important|must|shall|summary|conclusion|total|due|covered|"
                       r"excluded|exclusion|benefit|liable|responsible|deadline|penalty|"
                       r"agree|purpose|objective|result|recommend|key|period|premium|payable)\b")


def split_sentences(text):
    text = re.sub(r"-\n(?=[a-z])", "", text)
    text = re.sub(r"\s+", " ", text)
    # Don't end a sentence at abbreviations like "Rs." or "No."
    text = re.sub(r"(?i)\b(rs|no|nos|mr|mrs|ms|dr|sr|jr|st|co|ltd|pvt|inc|vs|etc|approx|"
                  r"fig|vol|sec|art|cl|dept|govt|e\.g|i\.e)\.", r"\1<dot>", text)
    parts = [p.replace("<dot>", ".") for p in
             re.split(r"(?<=[.!?])\s+(?=[A-Z0-9\"'(“])", text)]
    good = []
    for part in parts:
        words = part.split()
        if not 6 <= len(words) <= 60:
            continue
        letters = sum(c.isalpha() for c in part)
        if letters < 0.55 * len(part):
            continue
        good.append(part.strip())
    return good


def summarize_text(text, word_budget=260):
    sentences = split_sentences(text)
    tokens_per = [re.findall(r"[a-z][a-z'-]{2,}", s.lower()) for s in sentences]
    freq = Counter(t for tokens in tokens_per for t in tokens if t not in STOPWORDS)
    if not sentences or not freq:
        return [], []
    top = max(freq.values())
    scored = []
    for index, (sentence, tokens) in enumerate(zip(sentences, tokens_per)):
        content = [t for t in tokens if t not in STOPWORDS]
        if not content:
            continue
        score = sum(freq[t] / top for t in content) / (len(content) ** 0.7)
        if index < max(3, len(sentences) // 10):
            score *= 1.25                      # opening sentences matter
        if CUE_WORDS.search(sentence):
            score *= 1.15
        scored.append((score, index, set(content)))

    # Short documents: keep about half of the distinct sentences
    limit = min(12, max(3, (len(set(sentences)) + 1) // 2))
    chosen, used_words = [], 0
    for score, index, words in sorted(scored, reverse=True):
        if any(len(words & other) / max(1, len(words | other)) > 0.5 for _, other in chosen):
            continue                           # skip near-duplicates
        length = len(sentences[index].split())
        if used_words + length > word_budget and chosen:
            continue
        chosen.append((index, words))
        used_words += length
        if used_words >= word_budget or len(chosen) >= limit:
            break
    summary = [sentences[i] for i, _ in sorted(chosen)]
    key_terms = [w for w, _ in freq.most_common(40) if len(w) > 3 and not w.isdigit()][:12]
    return summary, key_terms


def unique_first(items, limit):
    seen, out = set(), []
    for item in items:
        key = item.lower().strip()
        if key and key not in seen:
            seen.add(key)
            out.append(item.strip())
        if len(out) >= limit:
            break
    return out


def guess_title(doc, first_text, path):
    meta_title = (doc.metadata or {}).get("title", "") or ""
    if len(meta_title) > 3 and not re.search(r"(?i)microsoft|\.docx?|untitled", meta_title):
        return meta_title.strip()
    for line in first_text.splitlines()[:15]:
        words = line.split()
        if (2 <= len(words) <= 14 and sum(c.isalpha() for c in line) > 0.6 * len(line)
                and not line.rstrip().endswith((".", ",")) and not line[:1].islower()):
            return line.strip()
    return os.path.splitext(os.path.basename(path))[0]


def build_summary(worker, path):
    doc = open_document(path)
    total = len(doc)
    texts, ocr_pages = [], 0
    for i, page in enumerate(doc):
        if worker.cancelled:
            return None
        worker.report(i, total, f"Reading page {i + 1} of {total}...")
        text, used_ocr = read_page_text(page)
        texts.append(text)
        ocr_pages += used_ocr
    worker.report(total, total, "Writing the summary...")
    full = "\n".join(texts)
    summary, key_terms = summarize_text(full)
    if not summary:   # e.g. forms and tables with few full sentences
        lines = [l.strip() for l in full.splitlines() if len(l.split()) >= 3]
        summary = unique_first(lines, 12)

    amounts = sorted(unique_first(MONEY_RE.findall(full), 30), key=money_value, reverse=True)
    facts = [
        ("Reference numbers", unique_first([f"{m.group(1).title()} No: {m.group(2)}"
                                            for m in REF_RE.finditer(full)], 5)),
        ("Organisations", unique_first(COMPANY_RE.findall(full), 5)),
        ("Dates mentioned", unique_first(DATE_RE.findall(full), 6)),
        ("Largest amounts", amounts[:5]),
        ("Emails", unique_first(EMAIL_RE.findall(full), 3)),
    ]
    return {
        "title": guess_title(doc, texts[0] if texts else "", path),
        "file": os.path.basename(path),
        "pages": total,
        "words": len(full.split()),
        "ocr_pages": ocr_pages,
        "summary": summary,
        "key_terms": key_terms,
        "facts": [(label, values) for label, values in facts if values],
    }


def summary_html(info):
    esc = html.escape
    when = datetime.date.today().strftime("%d %b %Y")
    parts = [
        f"<h1>{esc(info['title'])}</h1>",
        f"<p class='meta'>{esc(info['file'])} &nbsp;|&nbsp; {info['pages']} page(s) "
        f"&nbsp;|&nbsp; {info['words']:,} words &nbsp;|&nbsp; summarised {when}</p>",
        "<h2>Summary</h2><ul>",
    ]
    parts += [f"<li>{esc(s)}</li>" for s in info["summary"]] or ["<li>No readable text found.</li>"]
    parts.append("</ul>")
    if info["facts"]:
        parts.append("<h2>Key facts</h2><table>")
        for label, values in info["facts"]:
            parts.append(f"<tr><td class='lbl'>{esc(label)}</td>"
                         f"<td>{esc(';  '.join(values))}</td></tr>")
        parts.append("</table>")
    if info["key_terms"]:
        parts.append("<h2>Key terms</h2><p>" + esc("  ·  ".join(info["key_terms"])) + "</p>")
    note = ("Automatic summary: the most important sentences were selected from the "
            "document. Please check against the original for important decisions.")
    if info["ocr_pages"]:
        note += f" {info['ocr_pages']} scanned page(s) were read with OCR."
    parts.append(f"<p class='note'>{esc(note)}</p>")
    return "\n".join(parts)


SUMMARY_CSS = """
* { font-family: sans-serif; }
h1 { color: #1d976c; font-size: 20px; margin: 0 0 4px 0; }
h2 { color: #1565c0; font-size: 14px; margin: 12px 0 4px 0; border-bottom: 1px solid #ccc; }
p, li, td { font-size: 11px; color: #222; line-height: 1.35; }
.meta { color: #666; font-size: 10px; }
.note { color: #888; font-size: 9px; margin-top: 12px; }
td.lbl { font-weight: bold; color: #333; width: 130px; vertical-align: top; }
"""


def save_summary_pdf(info, out_path):
    """Write the summary on exactly one A4 page (text shrinks to fit if needed)."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.draw_rect(fitz.Rect(0, 0, 595, 8), color=None, fill=(0.11, 0.59, 0.42))
    page.insert_htmlbox(fitz.Rect(42, 34, 553, 812), summary_html(info),
                        css=SUMMARY_CSS, scale_low=0)
    doc.set_metadata({"title": f"Summary - {info['title']}", "creator": APP_NAME})
    doc.save(out_path, garbage=3, deflate=True)


class SummaryPage(QWidget):
    def __init__(self):
        super().__init__()
        self.path = None
        self.info = None
        self.worker = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)
        layout.addWidget(header_banner(
            "📝 1-Page Summary",
            "Get the key points of a long document on a single page. "
            "Scanned pages are read with OCR.", SUMMARY_COLORS))

        card = white_card()
        col = QVBoxLayout(card)
        col.setContentsMargins(18, 14, 18, 14)
        col.setSpacing(8)
        row = QHBoxLayout()
        choose = gradient_button("📂 Choose File", BLUE)
        choose.clicked.connect(self.choose_file)
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("Info")
        self.run_btn = gradient_button("📝 Create Summary", SUMMARY_COLORS)
        self.run_btn.clicked.connect(self.run)
        self.run_btn.setEnabled(False)
        row.addWidget(choose)
        row.addSpacing(8)
        row.addWidget(self.file_label, 1)
        row.addWidget(self.run_btn)
        col.addLayout(row)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        col.addWidget(self.progress)

        self.view = QTextBrowser()
        self.view.setOpenExternalLinks(False)
        self.view.setPlaceholderText("Your one-page summary will appear here.")
        self.view.document().setDefaultStyleSheet(SUMMARY_CSS.replace("px", "pt"))
        col.addWidget(self.view, 1)

        out = QHBoxLayout()
        out.addStretch()
        self.copy_btn = gradient_button("📋 Copy Text", PURPLE)
        self.pdf_btn = gradient_button("💾 Save 1-Page PDF", GREEN)
        self.copy_btn.clicked.connect(self.copy_text)
        self.pdf_btn.clicked.connect(self.save_pdf)
        out.addWidget(self.copy_btn)
        out.addWidget(self.pdf_btn)
        col.addLayout(out)
        layout.addWidget(card, 1)
        self.set_result_buttons(False)
        if not HAS_PREVIEWS:
            self.view.setPlainText("This tool needs PyMuPDF:  pip install pymupdf")
            choose.setEnabled(False)

    def set_result_buttons(self, enabled):
        self.copy_btn.setEnabled(enabled)
        self.pdf_btn.setEnabled(enabled)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a document", "",
            "PDF or image (*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp)")
        if not path:
            return
        self.path = path
        self.file_label.setText(os.path.basename(path))
        self.file_label.setToolTip(path)
        self.run_btn.setEnabled(True)

    def run(self):
        self.run_btn.setEnabled(False)
        self.set_result_buttons(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.view.clear()
        self.worker = TaskWorker(build_summary, self.path)
        self.worker.progress.connect(self.on_progress)
        self.worker.done.connect(self.on_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_progress(self, done, total, message):
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(done)
        set_status(self, message)

    def on_done(self, info):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        if info is None:
            return
        self.info = info
        self.view.setHtml(summary_html(info))
        self.set_result_buttons(True)
        set_status(self, "Summary ready.")

    def on_failed(self, message):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        show_error(self, f"Could not summarise this file:\n\n{message}")

    def copy_text(self):
        QApplication.clipboard().setText(self.view.toPlainText())
        set_status(self, "Summary copied to clipboard.")

    def save_pdf(self):
        base = os.path.splitext(os.path.basename(self.path))[0]
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save summary as", os.path.join(os.path.dirname(self.path),
                                                  f"{base}_summary.pdf"), "PDF files (*.pdf)")
        if not out_path:
            return
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        if same_file(out_path, self.path):
            show_error(self, "Please choose a different name than the original file.")
            return
        try:
            save_summary_pdf(self.info, out_path)
        except Exception as err:
            show_error(self, f"Saving failed:\n\n{err}")
            return
        show_success(self, f"One-page summary saved to:\n\n{out_path}",
                     os.path.dirname(out_path))


# ======================================================================
#  Invoices -> Excel
# ======================================================================

INVOICE_WORD_RE = re.compile(r"(?i)\b(tax\s*invoice|invoice|bill\s*of\s*supply|"
                             r"credit\s+note|debit\s+note)\b")
AMOUNT_RE = re.compile(r"(?<![\w.,/-])(\d{1,3}(?:,\d{2,3})+(?:\.\d{1,2})?|\d+\.\d{1,2}|\d{2,})"
                       r"(?![\w%/-]|[.,]\d)")
TAX_STOP_RE = re.compile(r"(?i)\b(cgst|sgst|utgst|igst|cess|total|round)")

INVOICE_NO_LABELS = [r"(?:tax\s+)?invoice\s*(?:no|number|num|#)\.?", r"\binv\.?\s*(?:no|#)\.?",
                     r"bill\s*(?:no|number)\.?", r"(?:credit|debit)\s*note\s*(?:no|number)\.?",
                     r"(?:document|doc)\s*no\.?"]
INVOICE_DATE_LABELS = [r"invoice\s*date", r"inv\.?\s*date", r"bill\s*date",
                       r"date\s*of\s*(?:invoice|issue)", r"\bdated\b", r"\bdate\b"]
TAXABLE_LABELS = [r"taxable\s*(?:value|amount|amt)", r"assessable\s*value",
                  r"sub\s*-?\s*total", r"total\s*(?:before|excl\.?|excluding)\s*tax"]
TOTAL_LABELS = [r"grand\s*total", r"total\s*invoice\s*(?:value|amount)",
                r"invoice\s*(?:value|total|amount)", r"total\s*amount\s*(?:payable|due|after\s*tax)?",
                r"(?:net|balance)\s*(?:amount|payable)", r"amount\s*(?:payable|due)",
                r"^\s*total\b"]
BUYER_RE = re.compile(r"(?i)bill(?:ed)?\s*to|buyer|recipient|consignee|ship(?:ped)?\s*to|"
                      r"customer|sold\s*to|party\s*(?:name|details)")
SELLER_RE = re.compile(r"(?i)supplier|seller|sold\s*by|consignor|vendor|billed\s*by")
DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %b %y", "%d %B %Y", "%d-%B-%Y",
                "%b %d %Y", "%B %d %Y", "%Y-%m-%d"]

INVOICE_COLUMNS = ["File", "Pages", "Type", "Invoice No", "Invoice Date", "Supplier",
                   "Supplier GSTIN", "Buyer", "Buyer GSTIN", "Taxable Value", "CGST",
                   "SGST/UTGST", "IGST", "Invoice Value", "Check"]
NUMBER_COLUMNS = range(9, 14)    # Taxable Value .. Invoice Value


def to_amount(text):
    try:
        return float(text.replace(",", ""))
    except (TypeError, ValueError):
        return None


def amounts_in(text):
    values = []
    for m in AMOUNT_RE.finditer(text):
        if text[m.end():m.end() + 2].strip().startswith("%"):
            continue
        value = to_amount(m.group(1))
        if value is not None:
            values.append(value)
    return values


def find_amount(lines, patterns, last=True):
    for pattern in patterns:
        found = []
        for idx, line in enumerate(lines):
            m = re.search(pattern, line, re.I)
            if not m or re.search(r"(?i)in\s+words|rupees\s+\w+.*only", line):
                continue
            rest = line[m.end():]
            stop = TAX_STOP_RE.search(rest)
            values = amounts_in(rest[:stop.start()] if stop else rest)
            if not values and idx + 1 < len(lines) and not re.search(r"[A-Za-z]{4,}", lines[idx + 1]):
                values = amounts_in(lines[idx + 1])
            if values:
                found.append(values[-1])
        if found:
            return found[-1] if last else found[0]
    return None


def find_text_value(lines, patterns, value_re, search=False):
    for pattern in patterns:
        for idx, line in enumerate(lines):
            m = re.search(pattern, line, re.I)
            if not m:
                continue
            candidates = [line[m.end():].lstrip(" :.-#\t")]
            if idx + 1 < len(lines):
                candidates.append(lines[idx + 1].strip())
            for text in candidates:
                vm = value_re.search(text[:45]) if search else value_re.match(text)
                if vm and re.search(r"\d", vm.group(0)):
                    return vm.group(0).strip(" .,:")
    return None


def parse_date(text):
    if not text:
        return None
    clean = re.sub(r"(?<=\d)(st|nd|rd|th)", "", text)
    clean = re.sub(r"[,]", " ", clean)
    clean = re.sub(r"\s+", " ", clean).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(clean, fmt).date()
        except ValueError:
            continue
    return None


def gstins_by_role(lines):
    supplier = buyer = None
    for idx, line in enumerate(lines):
        for gstin in GSTIN_RE.findall(line):
            role = None
            for back in range(idx, max(-1, idx - 9), -1):
                if BUYER_RE.search(lines[back]):
                    role = "buyer"
                    break
                if SELLER_RE.search(lines[back]):
                    role = "supplier"
                    break
            if role == "buyer" and buyer is None and gstin != supplier:
                buyer = gstin
            elif role != "buyer" and supplier is None:
                supplier = gstin
            elif buyer is None and gstin != supplier:
                buyer = gstin
    return supplier, buyer


def party_names(lines):
    """Supplier = name at the top (before the buyer block); buyer = after 'Bill To' etc."""
    skip = re.compile(r"(?i)tax\s*invoice|^invoice|original|duplicate|triplicate|gstin|"
                      r"page\s*\d|^\W*$")
    buyer_at = next((i for i, l in enumerate(lines[:40]) if BUYER_RE.search(l)), None)
    buyer = None
    if buyer_at is not None:
        line = lines[buyer_at]
        rest = BUYER_RE.split(line, maxsplit=1)[-1].strip(" :.-\t")
        if len(re.sub(r"[^A-Za-z]", "", rest)) >= 3:
            buyer = rest[:60]
        elif buyer_at + 1 < len(lines):
            buyer = lines[buyer_at + 1][:60]
    top = lines[:buyer_at] if buyer_at else lines[:12]
    supplier = None
    for line in top:
        m = COMPANY_RE.search(line)
        if m and not skip.search(line):
            supplier = m.group(0)
            break
    if supplier is None:
        for line in top[:8]:
            if not skip.search(line) and len(re.findall(r"[A-Za-z]{2,}", line)) >= 2 \
                    and not re.search(r"(?i)\b(no|date|dated)\b", line):
                supplier = line[:60]
                break
    return supplier, buyer


def extract_invoice(lines):
    type_match = INVOICE_WORD_RE.search(" ".join(lines[:30]))
    doc_type = "Invoice"
    if type_match:
        doc_type = re.sub(r"(?i)^(tax|bill|credit|debit)\s*", r"\1 ", type_match.group(1))
        doc_type = re.sub(r"(?i)\s*of\s*supply", " of Supply", re.sub(r"\s+", " ", doc_type)).title()
        doc_type = doc_type.replace("Of Supply", "of Supply")
    date_text = find_text_value(lines, INVOICE_DATE_LABELS, DATE_RE, search=True)
    supplier_gstin, buyer_gstin = gstins_by_role(lines)
    supplier, buyer = party_names(lines)
    return {
        "Type": doc_type,
        "Invoice Date": parse_date(date_text) or date_text,
        "Supplier": supplier, "Supplier GSTIN": supplier_gstin,
        "Buyer": buyer, "Buyer GSTIN": buyer_gstin,
        "Taxable Value": find_amount(lines, TAXABLE_LABELS, last=False),
        "CGST": find_amount(lines, [r"\bcgst\b"]),
        "SGST/UTGST": find_amount(lines, [r"\b(?:sgst|utgst)\b"]),
        "IGST": find_amount(lines, [r"\bigst\b"]),
        "Invoice Value": find_amount(lines, TOTAL_LABELS, last=True),
    }


def read_invoices(worker, paths):
    rows = []
    pages_total = 0
    docs = []
    for path in paths:
        doc = open_document(path)
        docs.append((path, doc))
        pages_total += len(doc)
    done = 0
    for path, doc in docs:
        current = None
        for i, page in enumerate(doc):
            if worker.cancelled:
                return None
            done += 1
            worker.report(done, pages_total,
                          f"Reading {os.path.basename(path)} - page {i + 1}")
            text, _ = read_page_text(page)
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            is_invoice = INVOICE_WORD_RE.search(text) is not None
            number = find_text_value(lines, INVOICE_NO_LABELS,
                                     re.compile(r"[A-Z0-9][A-Z0-9/\-_.]{0,29}", re.I)) \
                if is_invoice else None
            if number:
                number = re.sub(r"(?i)(invoice|inv|date|dated|dt)+\.?$", "", number).strip(" -/.")
            if number and (current is None or number != current["Invoice No"]):
                current = {"File": os.path.basename(path), "pages": [i + 1],
                           "Invoice No": number, "lines": list(lines)}
                rows.append(current)
            elif current is None and is_invoice:
                current = {"File": os.path.basename(path), "pages": [i + 1],
                           "Invoice No": None, "lines": list(lines)}
                rows.append(current)
            elif current is not None:
                current["pages"].append(i + 1)
                current["lines"] += lines
    result = []
    for row in rows:
        data = extract_invoice(row["lines"])
        data["File"] = row["File"]
        data["Pages"] = describe_pages(row["pages"])
        data["Invoice No"] = row["Invoice No"]
        result.append(data)
    return result


def invoice_check(values):
    total = values.get("Invoice Value")
    if total is None:
        return "Value missing"
    parts = sum(values.get(k) or 0 for k in ("Taxable Value", "CGST", "SGST/UTGST", "IGST"))
    diff = total - parts
    return "OK" if abs(diff) <= 1 else f"Diff {diff:,.2f}"


def export_invoices_xlsx(rows, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    base = Font(name="Arial", size=10)
    bold_white = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1565C0")
    total_fill = PatternFill("solid", start_color="E3F2FD")
    thin = Side(style="thin", color="D0D7E5")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    indian = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00'

    for col, name in enumerate(INVOICE_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font, cell.fill, cell.border = bold_white, head_fill, border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, values in enumerate(rows, start=2):
        for col, name in enumerate(INVOICE_COLUMNS, start=1):
            cell = ws.cell(row=r, column=col)
            cell.font, cell.border = base, border
            if name == "Check":
                cell.value = (f'=IF(N{r}="","Value missing",IF(ABS(N{r}-SUM(J{r}:M{r}))<=1,'
                              f'"OK","Diff "&TEXT(N{r}-SUM(J{r}:M{r}),"#,##0.00")))')
                continue
            value = values.get(name)
            cell.value = value if value not in ("",) else None
            if col - 1 in NUMBER_COLUMNS:
                cell.number_format = indian
            elif name == "Invoice Date" and isinstance(value, datetime.date):
                cell.number_format = "dd-mmm-yyyy"

    last = len(rows) + 1
    total_row = last + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(name="Arial", bold=True)
    for col in range(1, len(INVOICE_COLUMNS) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.fill, cell.border = total_fill, border
        if col - 1 in NUMBER_COLUMNS:
            letter = get_column_letter(col)
            cell.value = f"=SUM({letter}2:{letter}{last})" if rows else 0
            cell.number_format = indian
            cell.font = Font(name="Arial", size=10, bold=True)

    widths = [24, 8, 13, 18, 13, 30, 18, 30, 18, 15, 13, 13, 13, 16, 16]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(INVOICE_COLUMNS))}{max(last, 1)}"

    notes = wb.create_sheet("About")
    lines = [
        ("Invoice register", True),
        (f"Created by {APP_NAME} on {datetime.date.today():%d-%b-%Y}.", False),
        ("Values were read automatically from the invoice files (OCR for scanned pages).", False),
        ("Please verify against the original invoices before posting or filing.", False),
        ("Dates are read as day/month/year (Indian format).", False),
        ("Check column: OK when Taxable Value + CGST + SGST/UTGST + IGST equals Invoice Value "
         "within 1.00; a difference may be cess, round-off or a value that was not found.", False),
    ]
    for r, (text, is_bold) in enumerate(lines, start=1):
        notes.cell(row=r, column=1, value=text).font = Font(name="Arial", size=10, bold=is_bold)
    notes.column_dimensions["A"].width = 110
    wb.save(out_path)


class InvoicePage(QWidget):
    def __init__(self):
        super().__init__()
        self.paths = []
        self.worker = None
        self.updating = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)
        layout.addWidget(header_banner(
            "🧾 Invoices to Excel",
            "Read invoice number, date, GSTINs, taxable value, taxes and invoice value "
            "into an Excel register. You can correct any cell before exporting.",
            INVOICE_COLORS))

        card = white_card()
        col = QVBoxLayout(card)
        col.setContentsMargins(18, 14, 18, 14)
        col.setSpacing(8)
        row = QHBoxLayout()
        choose = gradient_button("📂 Choose Invoice Files", BLUE)
        choose.clicked.connect(self.choose_files)
        self.file_label = QLabel("No files selected (you can pick several PDFs or images)")
        self.file_label.setObjectName("Info")
        self.read_btn = gradient_button("🔍 Read Invoices", INVOICE_COLORS)
        self.read_btn.clicked.connect(self.run)
        self.read_btn.setEnabled(False)
        row.addWidget(choose)
        row.addSpacing(8)
        row.addWidget(self.file_label, 1)
        row.addWidget(self.read_btn)
        col.addLayout(row)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        col.addWidget(self.progress)

        self.table = QTableWidget(0, len(INVOICE_COLUMNS))
        self.table.setHorizontalHeaderLabels(INVOICE_COLUMNS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.itemChanged.connect(self.on_item_changed)
        for c, w in enumerate([150, 64, 100, 130, 115, 190, 150, 180, 150, 115, 90, 105, 90, 120]):
            self.table.setColumnWidth(c, w)
        col.addWidget(self.table, 1)

        hint = QLabel("Tip: double-click any cell to correct it. 'Check' compares "
                      "Taxable + taxes with the Invoice Value.")
        hint.setObjectName("Hint")
        col.addWidget(hint)

        out = QHBoxLayout()
        out.addStretch()
        self.clear_btn = gradient_button("🧹 Clear", RED)
        self.export_btn = gradient_button("📊 Export to Excel", GREEN)
        self.clear_btn.clicked.connect(lambda: self.table.setRowCount(0))
        self.export_btn.clicked.connect(self.export)
        out.addWidget(self.clear_btn)
        out.addWidget(self.export_btn)
        col.addLayout(out)
        layout.addWidget(card, 1)
        if not HAS_PREVIEWS:
            choose.setEnabled(False)
            self.file_label.setText("This tool needs PyMuPDF:  pip install pymupdf")

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Choose invoice files", "",
            "PDF or image (*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp)")
        if not paths:
            return
        self.paths = paths
        names = ", ".join(os.path.basename(p) for p in paths[:3])
        more = f" + {len(paths) - 3} more" if len(paths) > 3 else ""
        self.file_label.setText(f"{len(paths)} file(s): {names}{more}")
        self.read_btn.setEnabled(True)

    def run(self):
        self.read_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.worker = TaskWorker(read_invoices, list(self.paths))
        self.worker.progress.connect(self.on_progress)
        self.worker.done.connect(self.on_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_progress(self, done, total, message):
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(done)
        set_status(self, message)

    def on_failed(self, message):
        self.progress.setVisible(False)
        self.read_btn.setEnabled(True)
        show_error(self, f"Could not read the invoices:\n\n{message}")

    def on_done(self, rows):
        self.progress.setVisible(False)
        self.read_btn.setEnabled(True)
        if rows is None:
            return
        self.updating = True
        for values in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            for c, name in enumerate(INVOICE_COLUMNS):
                value = values.get(name)
                if isinstance(value, datetime.date):
                    text = value.strftime("%d-%m-%Y")
                elif isinstance(value, float):
                    text = f"{value:,.2f}"
                else:
                    text = "" if value is None else str(value)
                item = QTableWidgetItem(text)
                if c in NUMBER_COLUMNS:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if name in ("File", "Pages", "Check"):
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(r, c, item)
            self.refresh_check(r)
        self.updating = False
        found = len(rows)
        set_status(self, f"Found {found} invoice(s).")
        if found == 0:
            show_error(self, "No invoices were recognised in these files.")

    def row_values(self, r):
        values = {}
        for c, name in enumerate(INVOICE_COLUMNS):
            item = self.table.item(r, c)
            text = item.text().strip() if item else ""
            if c in NUMBER_COLUMNS:
                values[name] = to_amount(text) if text else None
            elif name == "Invoice Date":
                values[name] = parse_date(text) or (text or None)
            elif name == "Check":
                continue
            else:
                values[name] = text or None
        return values

    def refresh_check(self, r):
        item = self.table.item(r, INVOICE_COLUMNS.index("Check"))
        if item is None:
            return
        result = invoice_check(self.row_values(r))
        was_updating, self.updating = self.updating, True
        item.setText(result)
        item.setForeground(QColor("#2e7d32" if result == "OK" else "#c62828"))
        self.updating = was_updating

    def on_item_changed(self, item):
        if not self.updating and item.column() != INVOICE_COLUMNS.index("Check"):
            self.refresh_check(item.row())

    def export(self):
        if not HAS_XLSX:
            show_error(self, "Excel export needs openpyxl. Run:\n\n    pip install openpyxl")
            return
        if self.table.rowCount() == 0:
            show_error(self, "There are no invoices to export yet.")
            return
        start_dir = os.path.dirname(self.paths[0]) if self.paths else ""
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel file as", os.path.join(start_dir, "invoice_register.xlsx"),
            "Excel files (*.xlsx)")
        if not out_path:
            return
        if not out_path.lower().endswith(".xlsx"):
            out_path += ".xlsx"
        try:
            export_invoices_xlsx([self.row_values(r) for r in range(self.table.rowCount())],
                                 out_path)
        except PermissionError:
            show_error(self, "Could not save. Is the file open in Excel? Close it and try again.")
            return
        except Exception as err:
            show_error(self, f"Export failed:\n\n{err}")
            return
        show_success(self, f"Excel register with {self.table.rowCount()} invoice(s) "
                           f"saved to:\n\n{out_path}", os.path.dirname(out_path))


# ======================================================================
#  Redaction - detection rules
# ======================================================================

def only_digits(text):
    return re.sub(r"\D", "", text)


def luhn_ok(digits):
    if not digits:
        return False
    total, double = 0, False
    for ch in reversed(digits):
        n = int(ch)
        if double:
            n = n * 2 - 9 if n > 4 else n * 2
        total += n
        double = not double
    return total % 10 == 0


_V_D = [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9], [1, 2, 3, 4, 0, 6, 7, 8, 9, 5],
        [2, 3, 4, 0, 1, 7, 8, 9, 5, 6], [3, 4, 0, 1, 2, 8, 9, 5, 6, 7],
        [4, 0, 1, 2, 3, 9, 5, 6, 7, 8], [5, 9, 8, 7, 6, 0, 4, 3, 2, 1],
        [6, 5, 9, 8, 7, 1, 0, 4, 3, 2], [7, 6, 5, 9, 8, 2, 1, 0, 4, 3],
        [8, 7, 6, 5, 9, 3, 2, 1, 0, 4], [9, 8, 7, 6, 5, 4, 3, 2, 1, 0]]
_V_P = [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9], [1, 5, 7, 6, 2, 8, 3, 0, 9, 4],
        [5, 8, 0, 3, 7, 9, 6, 1, 4, 2], [8, 9, 1, 6, 0, 4, 3, 5, 2, 7],
        [9, 4, 5, 3, 1, 2, 6, 8, 7, 0], [4, 2, 8, 6, 5, 7, 3, 9, 0, 1],
        [2, 7, 9, 3, 8, 0, 6, 4, 1, 5], [7, 0, 4, 6, 9, 1, 3, 2, 5, 8]]


def verhoeff_ok(digits):
    """Aadhaar numbers carry a Verhoeff check digit - this removes false matches."""
    check = 0
    for i, ch in enumerate(reversed(digits)):
        check = _V_D[check][_V_P[i % 8][int(ch)]]
    return check == 0


NAME_STOP = set("""policy details number no date address of the and plan premium insured
proposer sum amount name company limited ltd bank branch code type total period from to
status id card holder details signature place account email mobile phone gender age
relationship nominee dob birth""".split())


def clean_name(text):
    words = text.split()
    while words and words[-1].lower().strip(".") in NAME_STOP:
        words.pop()
    if not words or words[0].lower().strip(".") in NAME_STOP:
        return None
    name = " ".join(words)
    return name if len(name) >= 3 else None


def clean_card(text):
    d = only_digits(text)
    return text if 13 <= len(d) <= 19 and luhn_ok(d) else None


def clean_aadhaar(text):
    d = only_digits(text)
    return text if len(d) == 12 and verhoeff_ok(d) else None


def clean_sin(text):
    return text if luhn_ok(only_digits(text)) else None


def needs_digit(text):
    return text if re.search(r"\d", text) else None


def rule(category, pattern, group=0, checked=True, clean=None):
    return {"cat": category, "re": re.compile(pattern), "group": group,
            "on": checked, "clean": clean}


SMART_RULES = [
    rule("Name", r"\b(?:Mr|Mrs|Ms|Miss|Dr|Shri|Smt|Sri|Kumari|Prof|MR|MRS|MS|DR|SHRI|SMT)\.?[ ]+"
                 r"([A-Z][A-Za-z]+(?:[ ][A-Z]\.?(?![A-Za-z]))?(?:[ ][A-Z][A-Za-z]+){0,3})", 1, True, clean_name),
    rule("Name", r"(?i:\b(?:name|s/o|d/o|w/o|c/o|son\s+of|daughter\s+of|wife\s+of|insured|"
                 r"proposer|nominee|employee|applicant|patient|policy\s*holder)\b)[ ]*[:\-]?[ ]*"
                 r"([A-Z][A-Za-z.]+(?:[ ][A-Z][A-Za-z.]+){1,3})", 1, True, clean_name),
    rule("Email", EMAIL_RE.pattern),
    rule("Phone (India)", r"(?<![\d+])(?:\+91[ -]?|0)?[6-9]\d{4}[ -]?\d{5}(?!\d)"),
    rule("Phone (US)", r"(?<![\d+])(?:\+1[ .-]?)?\(?[2-9]\d{2}\)?[ .-]\d{3}[ .-]\d{4}(?!\d)"),
    rule("Phone (UK)", r"(?<![\d+])(?:\+44[ ]?|0)(?:7\d{3}[ ]?\d{3}[ ]?\d{3}|"
                       r"[1-3]\d{2,3}[ ]?\d{3}[ ]?\d{3,4})(?!\d)"),
    rule("PAN", r"\b[A-Z]{3}[ABCFGHLJPTK][A-Z]\d{4}[A-Z]\b"),
    rule("Aadhaar", r"(?<!\d)[2-9]\d{3}[ -]?\d{4}[ -]?\d{4}(?!\d)", 0, True, clean_aadhaar),
    rule("Passport", r"(?i:passport)[^\n]{0,25}?\b([A-PR-WY][1-9]\d{6})\b", 1),
    rule("Voter ID", r"\b[A-Z]{3}\d{7}\b"),
    rule("Driving Licence", r"\b[A-Z]{2}[- ]?\d{2}[- ]?(?:19|20)\d{2}[- ]?\d{7}\b"),
    rule("GSTIN", GSTIN_RE.pattern),
    rule("CIN", r"\b[LU]\d{5}[A-Z]{2}\d{4}[A-Z]{3}\d{6}\b"),
    rule("TAN", r"\b[A-Z]{4}\d{5}[A-Z]\b"),
    rule("DIN", r"(?i:\bDIN)\s*(?:No\.?|Number)?\s*[:\-]?\s*(\d{8})\b", 1),
    rule("LLPIN", r"(?i:LLPIN)\s*[:\-]?\s*([A-Z]{3}-\d{4})\b", 1),
    rule("FSSAI", r"(?i:FSSAI)[^\n\d]{0,30}(\d{14})\b", 1),
    rule("IFSC", r"\b[A-Z]{4}0[A-Z0-9]{6}\b"),
    rule("Bank Account", r"(?i:a/?c|account|acct)\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*(\d{9,18})\b", 1),
    rule("Credit/Debit Card", r"(?<!\d)(?:\d{4}[ -]?){3}\d{1,7}(?!\d)", 0, True, clean_card),
    rule("US SSN", r"\b(?!000|666|9\d\d)\d{3}-(?!00)\d{2}-(?!0000)\d{4}\b"),
    rule("UK National Insurance", r"\b(?!BG|GB|NK|KN|TN|NT|ZZ)[A-CEGHJ-PR-TW-Z][A-CEGHJ-NPR-TW-Z]"
                                  r"[ ]?\d{2}[ ]?\d{2}[ ]?\d{2}[ ]?[A-D]\b"),
    rule("Canadian SIN", r"(?<!\d)\d{3}[ -]\d{3}[ -]\d{3}(?!\d)", 0, True, clean_sin),
    rule("Date of Birth", r"(?i:d\.?o\.?b\.?|date\s+of\s+birth|birth\s*date|born(?:\s+on)?)"
                          r"\s*[:\-]?\s*(" + DATE_PATTERN + r")", 1),
    rule("Date", r"(?<![\d/.-])" + DATE_PATTERN + r"(?![\d/])", 0, False),
    rule("Company", COMPANY_RE.pattern),
    rule("Medical Record No.", r"(?i:\bMRN|medical\s+record\s*(?:no\.?|number)|UHID|patient\s*id|"
                               r"hospital\s*no\.?)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9/-]{3,19})",
         1, True, needs_digit),
    rule("Salary", r"(?i:salary|ctc|gross\s+pay|net\s+pay|basic\s+pay|take\s+home)[^\n\d]{0,20}"
                   r"((?:₹|Rs\.?|INR)?[ ]?\d[\d,]*(?:\.\d{1,2})?)", 1),
    rule("Money", MONEY_RE.pattern, 0, False),
    rule("Website", r"\b(?:https?://|www\.)[^\s<>\"')]+"),
    rule("IP Address", r"\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b"),
    rule("Signature (digital)", r"(?i:digitally\s+signed\s+by)[ :]*([A-Z][A-Za-z. ]{2,40})", 1),
]
CATEGORY_ORDER = {r["cat"]: i for i, r in enumerate(SMART_RULES)}
CATEGORY_ORDER.update({"Address (PIN code)": 26, "Signature (image)": 99, "Custom word": 100})

PIN_RE = re.compile(r"(?<!\d)[1-9]\d{2}[ ]?\d{3}(?!\d)")
ADDRESS_WORDS = re.compile(
    r"(?i)\b(road|rd\.?|street|st\.|nagar|colony|sector|lane|floor|building|bldg|tower|"
    r"dist(?:rict)?|village|post|p\.?o\.?|near|opp\.?|marg|block|plot|flat|apartment|apt|"
    r"layout|phase|cross|main|society|chowk|bazar|bazaar|city|town|state|pin|pincode|"
    r"west bengal|maharashtra|karnataka|tamil nadu|telangana|andhra pradesh|gujarat|delhi|"
    r"kolkata|mumbai|chennai|hyderabad|bengaluru|bangalore|pune|india)\b")


def detect_addresses(text):
    found = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if not PIN_RE.search(line) or MONEY_RE.search(line):
            continue
        prev_is_address = i > 0 and ADDRESS_WORDS.search(lines[i - 1])
        if ADDRESS_WORDS.search(line) or prev_is_address:
            found.append(("Address (PIN code)", line.strip(), True))
            if prev_is_address and len(lines[i - 1].strip()) > 3:
                found.append(("Address (PIN code)", lines[i - 1].strip(), True))
    return found


def detect_sensitive(text):
    """Return a list of (category, text, checked_by_default)."""
    found = []
    for r in SMART_RULES:
        for m in r["re"].finditer(text):
            value = (m.group(r["group"]) or "").strip()
            if r["clean"] and value:
                value = r["clean"](value)
            if value and len(value) >= 3:
                found.append((r["cat"], value, r["on"]))
    return found + detect_addresses(text)


def ocr_lines_for_page(page):
    """OCR one page -> [(line_text, rect in displayed page coordinates)]."""
    return [(text, rect) for rect, text in ocr_page(get_ocr_engine(), page)]


def locate_text(page, value, ocr_lines=None):
    """All rectangles (unrotated page coordinates) where 'value' appears."""
    if ocr_lines is None:
        return list(page.search_for(value))
    rects, needle = [], value.lower()
    for text, rect in ocr_lines:
        low, n, start = text.lower(), max(len(text), 1), 0
        while True:
            k = low.find(needle, start)
            if k < 0:
                break
            x0 = rect.x0 + rect.width * k / n
            x1 = rect.x0 + rect.width * (k + len(needle)) / n
            rects.append(fitz.Rect(x0, rect.y0, x1, rect.y1) * page.derotation_matrix)
            start = k + len(needle)
    return rects


def signature_images(page):
    """Small images placed close to 'Signature' / 'Signatory' / 'Signed' labels."""
    labels = []
    for word in ("Signature", "Signatory", "Signed", "Sign"):
        labels += page.search_for(word)
    if not labels:
        return []
    found = []
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        if r.is_empty or r.width > 300 or r.height > 200:
            continue                    # full-page scans and large pictures are not signatures
        for lab in labels:
            near_v = lab.y0 - 150 <= r.y1 <= lab.y1 + 40 or r.y0 - 40 <= lab.y0 <= r.y1 + 40
            near_h = r.x1 >= lab.x0 - 150 and r.x0 <= lab.x1 + 150
            if near_v and near_h:
                found.append(r)
                break
    return found


def analyze_for_redaction(worker, path, use_ocr):
    doc = open_document(path)
    total = len(doc)
    findings, ocr_cache, searched = {}, {}, set()
    for i in range(total):
        if worker.cancelled:
            return None
        worker.report(i, total, f"Analyzing page {i + 1} of {total}...")
        page = doc[i]
        text = page.get_text("text")
        lines = None
        if len(text.strip()) < 20 and use_ocr:
            lines = ocr_lines_for_page(page)
            ocr_cache[i] = lines
            text = boxes_to_text([(r, t) for t, r in lines])
        for category, value, checked in detect_sensitive(text):
            key = (category, value)
            item = findings.setdefault(key, {"cat": category, "text": value,
                                             "on": checked, "hits": []})
            if (key, i) in searched:
                continue
            searched.add((key, i))
            item["hits"] += [(i, r) for r in locate_text(page, value, lines)]
        for k, rect in enumerate(signature_images(page), start=1):
            findings[("Signature (image)", f"Signature image {k} on page {i + 1}")] = {
                "cat": "Signature (image)", "text": f"Signature image {k} on page {i + 1}",
                "on": True, "hits": [(i, rect)]}
    worker.report(total, total, "Done")
    result = [f for f in findings.values() if f["hits"]]
    result.sort(key=lambda f: (CATEGORY_ORDER.get(f["cat"], 50), f["hits"][0][0], f["text"]))
    return {"findings": result, "ocr": ocr_cache}


# ======================================================================
#  Redaction - document model and page canvas
# ======================================================================

class RedactionDocument:
    """Holds the open PDF and the redaction boxes (unrotated page coordinates)."""

    def __init__(self, path):
        self.path = path
        self.doc = open_document(path)
        self.boxes = {i: [] for i in range(len(self.doc))}
        self.ocr = {}      # page index -> OCR lines, filled by Analyze

    def __len__(self):
        return len(self.doc)

    def total_boxes(self):
        return sum(len(v) for v in self.boxes.values())

    def page_is_scanned(self, i):
        return len(self.doc[i].get_text("text").strip()) < 20

    def find(self, value):
        """{page: [rects]} for a word or phrase; also returns skipped scanned pages."""
        hits, skipped = {}, 0
        for i in range(len(self.doc)):
            page = self.doc[i]
            if i in self.ocr:
                rects = locate_text(page, value, self.ocr[i])
            elif self.page_is_scanned(i):
                skipped += 1
                continue
            else:
                rects = locate_text(page, value)
            if rects:
                hits[i] = rects
        return hits, skipped

    def redacted_page_pixmap(self, i, zoom):
        """Render page i exactly as it will look after saving."""
        tmp = fitz.open()
        tmp.insert_pdf(self.doc, from_page=i, to_page=i)
        page = tmp[0]
        for rect in self.boxes[i]:
            page.add_redact_annot(rect, fill=(0, 0, 0))
        if self.boxes[i]:
            page.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_PIXELS", 2))
        return page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)

    def save(self, out_path):
        out = fitz.open()
        out.insert_pdf(self.doc)
        for i, rects in self.boxes.items():
            if not rects:
                continue
            page = out[i]
            for rect in rects:
                page.add_redact_annot(rect, fill=(0, 0, 0))
            # Removes the text and image pixels under the boxes - not just covers them
            page.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_PIXELS", 2))
        out.set_metadata({})
        try:
            out.del_xml_metadata()
        except Exception:
            pass
        out.save(out_path, garbage=4, deflate=True, clean=True)


def pixmap_to_qpixmap(pix, dpr=1.0):
    image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
    qpix = QPixmap.fromImage(image.copy())
    qpix.setDevicePixelRatio(dpr)
    return qpix


class PageCanvas(QWidget):
    """Shows one page. Click-and-drag draws solid black redaction boxes."""

    boxDrawn = pyqtSignal(object)        # fitz.Rect (unrotated page coordinates)
    boxRemove = pyqtSignal(int)          # index of the box to delete

    def __init__(self):
        super().__init__()
        self.pixmap = None
        self.zoom = 1.0
        self.rot = self.derot = None
        self.boxes = []
        self.highlights = []
        self.brush = 12                  # points
        self.drag_start = self.drag_now = None
        self.setCursor(Qt.CrossCursor)
        self.setMouseTracking(True)

    def set_page(self, pixmap, zoom, page):
        self.pixmap, self.zoom = pixmap, zoom
        self.rot, self.derot = page.rotation_matrix, page.derotation_matrix
        size = pixmap.size() / pixmap.devicePixelRatio()
        self.setFixedSize(size)
        self.update()

    def to_screen(self, rect):
        r = fitz.Rect(rect) * self.rot
        return QRectF(r.x0 * self.zoom, r.y0 * self.zoom, r.width * self.zoom, r.height * self.zoom)

    def paintEvent(self, event):
        if self.pixmap is None:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.drawPixmap(0, 0, self.pixmap)
        p.setPen(QPen(QColor("#e53935"), 2))
        p.setBrush(QColor(229, 57, 53, 45))
        for rect in self.highlights:
            p.drawRect(self.to_screen(rect))
        p.setPen(QPen(QColor("#ff5252"), 1))
        p.setBrush(QColor("black"))
        for rect in self.boxes:
            p.drawRect(self.to_screen(rect))
        if self.drag_start is not None:
            p.setPen(QPen(QColor("#ff9800"), 2, Qt.DashLine))
            p.setBrush(QColor(0, 0, 0, 170))
            p.drawRect(self.drag_rect())
        p.end()

    def drag_rect(self):
        r = QRectF(self.drag_start, self.drag_now).normalized()
        min_size = self.brush * self.zoom
        if r.width() < min_size:            # thin drag -> brush-sized bar
            c = r.center().x()
            r.setLeft(c - min_size / 2)
            r.setRight(c + min_size / 2)
        if r.height() < min_size:
            c = r.center().y()
            r.setTop(c - min_size / 2)
            r.setBottom(c + min_size / 2)
        return r

    def mousePressEvent(self, event):
        if self.pixmap is None:
            return
        if event.button() == Qt.LeftButton:
            self.drag_start = self.drag_now = QPointF(event.pos())
            self.update()
        elif event.button() == Qt.RightButton:
            for index in range(len(self.boxes) - 1, -1, -1):
                if self.to_screen(self.boxes[index]).contains(QPointF(event.pos())):
                    self.boxRemove.emit(index)
                    break

    def mouseMoveEvent(self, event):
        if self.drag_start is not None:
            self.drag_now = QPointF(event.pos())
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.LeftButton or self.drag_start is None:
            return
        self.drag_now = QPointF(event.pos())
        r = self.drag_rect()
        self.drag_start = None
        shown = fitz.Rect(r.left() / self.zoom, r.top() / self.zoom,
                          r.right() / self.zoom, r.bottom() / self.zoom)
        self.boxDrawn.emit(shown * self.derot)
        self.update()


class RedactionPage(QWidget):
    def __init__(self):
        super().__init__()
        self.rdoc = None
        self.page_index = 0
        self.zoom = None              # None = fit to width
        self.search_hits = {}
        self.findings = []
        self.undo_stack = []
        self.unsaved = False
        self.worker = None
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.timeout.connect(self.update_preview)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(8)
        layout.addWidget(header_banner(
            "⬛ Redaction",
            "Hide sensitive information permanently. Draw boxes, search words, or let "
            "Smart Redaction find IDs, phone numbers, bank details and more.", REDACT_COLORS))

        # ---- top bar ----
        top = QHBoxLayout()
        self.open_btn = gradient_button("📂 Open PDF", BLUE)
        self.open_btn.clicked.connect(self.open_pdf)
        self.file_label = QLabel("No PDF open")
        self.file_label.setObjectName("Info")
        self.save_btn = gradient_button("💾 Save Redacted PDF", GREEN)
        self.save_btn.clicked.connect(self.save_pdf)
        top.addWidget(self.open_btn)
        top.addSpacing(8)
        top.addWidget(self.file_label, 1)
        top.addWidget(self.save_btn)
        layout.addLayout(top)

        body = QHBoxLayout()
        body.setSpacing(10)
        self.tabs = QTabWidget()
        self.tabs.addTab(self.build_manual_tab(), "✏️ Manual Redaction")
        self.tabs.addTab(self.build_smart_tab(), "🧠 Smart Redaction")
        body.addWidget(self.tabs, 1)

        preview = QFrame()
        preview.setObjectName("PreviewCard")
        preview.setFixedWidth(290)
        pv = QVBoxLayout(preview)
        pv.setContentsMargins(10, 8, 10, 10)
        title = QLabel("👁 Live Preview (as it will be saved)")
        title.setObjectName("PreviewTitle")
        title.setWordWrap(True)
        self.preview_label = QLabel("Open a PDF to see the preview")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setObjectName("Hint")
        self.preview_label.setMinimumHeight(200)
        self.preview_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.box_count = QLabel("")
        self.box_count.setObjectName("Hint")
        self.box_count.setWordWrap(True)
        pv.addWidget(title)
        pv.addWidget(self.preview_label, 1)
        pv.addWidget(self.box_count)
        body.addWidget(preview)
        layout.addLayout(body, 1)

    # ------------------------------------------------------------------
    def build_manual_tab(self):
        tab = QWidget()
        col = QVBoxLayout(tab)
        col.setContentsMargins(10, 10, 10, 10)
        col.setSpacing(6)

        nav = QHBoxLayout()
        prev_btn = gradient_button("◀ Previous", BLUE, compact=True)
        next_btn = gradient_button("Next ▶", BLUE, compact=True)
        prev_btn.clicked.connect(lambda: self.go_to(self.page_index - 1))
        next_btn.clicked.connect(lambda: self.go_to(self.page_index + 1))
        self.page_label = QLabel("Page - / -")
        self.page_label.setObjectName("Info")
        zoom_out = gradient_button("➖", PURPLE, compact=True)
        zoom_in = gradient_button("➕", PURPLE, compact=True)
        fit = gradient_button("Fit", PURPLE, compact=True)
        zoom_out.clicked.connect(lambda: self.change_zoom(1 / 1.25))
        zoom_in.clicked.connect(lambda: self.change_zoom(1.25))
        fit.clicked.connect(lambda: self.change_zoom(None))
        self.brush_label = QLabel("🖌 Brush: 12")
        self.brush_label.setObjectName("Hint")
        self.brush = QSlider(Qt.Horizontal)
        self.brush.setRange(4, 40)
        self.brush.setValue(12)
        self.brush.setFixedWidth(110)
        self.brush.setToolTip("Smallest box size. A quick drag along a line of text "
                              "makes a bar this thick; a single click makes a square.")
        self.brush.valueChanged.connect(self.brush_changed)
        for w in (prev_btn, self.page_label, next_btn):
            nav.addWidget(w)
        nav.addSpacing(10)
        for w in (zoom_out, zoom_in, fit):
            nav.addWidget(w)
        nav.addStretch()
        nav.addWidget(self.brush_label)
        nav.addWidget(self.brush)
        col.addLayout(nav)

        search = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Type a word or phrase to find…")
        self.search_edit.setMinimumHeight(36)
        self.search_edit.returnPressed.connect(self.do_search)
        find_btn = gradient_button("🔍 Find", BLUE, compact=True)
        find_btn.clicked.connect(self.do_search)
        redact_all = gradient_button("⬛ Redact All", ORANGE, compact=True)
        redact_all.setToolTip("Black out every match that was found")
        redact_all.clicked.connect(self.redact_search_hits)
        undo = gradient_button("↩ Undo", PURPLE, compact=True)
        undo.clicked.connect(self.undo)
        clear = gradient_button("🧹 Clear Page", RED, compact=True)
        clear.clicked.connect(self.clear_page)
        search.addWidget(self.search_edit, 1)
        for w in (find_btn, redact_all, undo, clear):
            search.addWidget(w)
        col.addLayout(search)

        self.canvas = PageCanvas()
        self.canvas.boxDrawn.connect(self.on_box_drawn)
        self.canvas.boxRemove.connect(self.on_box_remove)
        self.canvas_area = QScrollArea()
        self.canvas_area.setObjectName("CanvasArea")
        self.canvas_area.setWidget(self.canvas)
        self.canvas_area.setAlignment(Qt.AlignHCenter)
        self.canvas_area.setMinimumHeight(180)
        col.addWidget(self.canvas_area, 1)
        hint = QLabel("Drag to draw a black box  •  Right-click a box to remove it  •  "
                      "Red outlines = search matches")
        hint.setObjectName("Hint")
        col.addWidget(hint)
        return tab

    def build_smart_tab(self):
        tab = QWidget()
        col = QVBoxLayout(tab)
        col.setContentsMargins(10, 10, 10, 10)
        col.setSpacing(6)

        row = QHBoxLayout()
        self.analyze_btn = gradient_button("🧠 Analyze Document", BLUE)
        self.analyze_btn.clicked.connect(self.analyze)
        self.analyze_progress = QProgressBar()
        self.analyze_progress.setVisible(False)
        self.found_label = QLabel("Finds names, IDs, phones, bank details, addresses…")
        self.found_label.setObjectName("Hint")
        row.addWidget(self.analyze_btn)
        row.addWidget(self.analyze_progress, 1)
        row.addWidget(self.found_label, 1)
        col.addLayout(row)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["✔", "Category", "Found text", "Pages", "Hits"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setMinimumHeight(110)
        self.table.cellClicked.connect(self.on_finding_clicked)
        self.table.cellDoubleClicked.connect(
            lambda r, c: c and self.tabs.setCurrentIndex(0))
        col.addWidget(self.table, 1)

        sel = QHBoxLayout()
        all_btn = gradient_button("☑ Select All", PURPLE, compact=True)
        none_btn = gradient_button("☐ Deselect All", PURPLE, compact=True)
        tick_btn = gradient_button("✔ Tick Highlighted Rows", PURPLE, compact=True)
        all_btn.clicked.connect(lambda: self.set_all_checks(True))
        none_btn.clicked.connect(lambda: self.set_all_checks(False))
        tick_btn.clicked.connect(self.tick_selected_rows)
        apply_btn = gradient_button("✅ Apply Selected Redactions", GREEN, compact=True)
        apply_btn.clicked.connect(self.apply_selected)
        for w in (all_btn, none_btn, tick_btn):
            sel.addWidget(w)
        sel.addStretch()
        sel.addWidget(apply_btn)
        col.addLayout(sel)

        custom_label = QLabel("✍ Custom words to redact (comma-separated or one per line)")
        custom_label.setObjectName("Step")
        col.addWidget(custom_label)
        custom = QHBoxLayout()
        self.custom_edit = QPlainTextEdit()
        self.custom_edit.setPlaceholderText("e.g.  sohan, mohan, 112233E")
        self.custom_edit.setFixedHeight(64)
        custom.addWidget(self.custom_edit, 1)
        buttons = QVBoxLayout()
        buttons.setSpacing(4)
        add_btn = gradient_button("🔎 Find && Add to List", BLUE, compact=True)
        now_btn = gradient_button("⬛ Redact All Custom Words", ORANGE, compact=True)
        clr_btn = gradient_button("🧹 Clear", RED, compact=True)
        add_btn.clicked.connect(lambda: self.custom_words(add_to_list=True))
        now_btn.clicked.connect(lambda: self.custom_words(add_to_list=False))
        clr_btn.clicked.connect(self.custom_edit.clear)
        row2 = QHBoxLayout()
        row2.addWidget(add_btn)
        row2.addWidget(clr_btn)
        buttons.addLayout(row2)
        buttons.addWidget(now_btn)
        custom.addLayout(buttons)
        col.addLayout(custom)
        return tab

    # ------------------------------------------------------------------
    def require_doc(self):
        if self.rdoc is None:
            QMessageBox.information(self, APP_NAME, "⚠️ Please open a PDF first (📂 Open PDF).")
            return False
        return True

    def open_pdf(self):
        if self.unsaved and QMessageBox.question(
                self, APP_NAME, "You have redactions that are not saved yet.\n"
                                "Open another file and discard them?") != QMessageBox.Yes:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Open a PDF to redact", "", "PDF or image (*.pdf *.png *.jpg *.jpeg *.tif *.tiff)")
        if not path:
            return
        if not HAS_PREVIEWS:
            show_error(self, "Redaction needs PyMuPDF:  pip install pymupdf")
            return
        try:
            rdoc = RedactionDocument(path)
        except Exception as err:
            show_error(self, f"Could not open this file:\n\n{err}")
            return
        if len(rdoc) == 0:
            show_error(self, "This file has no pages.")
            return
        self.rdoc = rdoc
        self.search_hits, self.findings, self.undo_stack = {}, [], []
        self.table.setRowCount(0)
        self.unsaved = False
        scanned = sum(rdoc.page_is_scanned(i) for i in range(len(rdoc)))
        note = f"   |   {scanned} scanned page(s)" if scanned else ""
        self.file_label.setText(f"{os.path.basename(path)}   |   {len(rdoc)} page(s){note}")
        self.file_label.setToolTip(path)
        self.zoom = None
        self.go_to(0)
        if scanned:
            set_status(self, "Scanned pages found: run 🧠 Analyze Document so they can be "
                             "searched (OCR), or draw boxes by hand.")
        else:
            set_status(self, f"Opened {os.path.basename(path)}")

    def go_to(self, index):
        if self.rdoc is None:
            return
        index = max(0, min(index, len(self.rdoc) - 1))
        self.page_index = index
        self.page_label.setText(f"Page {index + 1} / {len(self.rdoc)}")
        self.render_page()

    def change_zoom(self, factor):
        if not self.require_doc():
            return
        if factor is None:
            self.zoom = None
        else:
            self.zoom = max(0.3, min(4.0, (self.zoom or self.fit_zoom()) * factor))
        self.render_page()

    def fit_zoom(self):
        page = self.rdoc.doc[self.page_index]
        width = max(200, self.canvas_area.viewport().width() - 24)
        return width / page.rect.width

    def render_page(self):
        page = self.rdoc.doc[self.page_index]
        zoom = self.zoom or self.fit_zoom()
        dpr = max(1.0, self.devicePixelRatioF())
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom * dpr, zoom * dpr), alpha=False)
        self.canvas.set_page(pixmap_to_qpixmap(pix, dpr), zoom, page)
        self.refresh_overlays()

    def refresh_overlays(self):
        if self.rdoc is None:
            return
        self.canvas.boxes = list(self.rdoc.boxes[self.page_index])
        self.canvas.highlights = list(self.search_hits.get(self.page_index, []))
        self.canvas.update()
        on_page = len(self.rdoc.boxes[self.page_index])
        self.box_count.setText(f"⬛ {on_page} box(es) on this page  •  "
                               f"{self.rdoc.total_boxes()} in total")
        self.preview_timer.start(250)

    def update_preview(self):
        if self.rdoc is None:
            return
        page = self.rdoc.doc[self.page_index]
        width = max(120, self.preview_label.width() - 6)
        height = max(120, self.preview_label.height() - 6)
        zoom = min(width / page.rect.width, height / page.rect.height)
        dpr = max(1.0, self.devicePixelRatioF())
        try:
            pix = self.rdoc.redacted_page_pixmap(self.page_index, zoom * dpr)
            self.preview_label.setPixmap(pixmap_to_qpixmap(pix, dpr))
        except Exception as err:
            self.preview_label.setText(f"Preview not available:\n{err}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.rdoc is not None and self.zoom is None:
            QTimer.singleShot(0, self.render_page)

    def brush_changed(self, value):
        self.brush_label.setText(f"🖌 Brush: {value}")
        self.canvas.brush = value

    # ---- adding / removing boxes ----
    def add_boxes(self, items, message):
        """items = [(page, rect)] - added as one undo step."""
        items = [(p, fitz.Rect(r)) for p, r in items if not fitz.Rect(r).is_empty]
        if not items:
            return 0
        for p, r in items:
            self.rdoc.boxes[p].append(r)
        self.undo_stack.append(("add", items))
        self.unsaved = True
        self.refresh_overlays()
        set_status(self, message.format(n=len(items)))
        return len(items)

    def on_box_drawn(self, rect):
        if self.rdoc is not None:
            self.add_boxes([(self.page_index, rect)], "Box added. Right-click it to remove.")

    def on_box_remove(self, index):
        rect = self.rdoc.boxes[self.page_index].pop(index)
        self.undo_stack.append(("remove", [(self.page_index, rect)]))
        self.unsaved = True
        self.refresh_overlays()
        set_status(self, "Box removed.")

    def undo(self):
        if not self.require_doc():
            return
        if not self.undo_stack:
            set_status(self, "Nothing to undo.")
            return
        action, items = self.undo_stack.pop()
        for p, r in items:
            if action == "add":
                if r in self.rdoc.boxes[p]:
                    self.rdoc.boxes[p].remove(r)
            else:
                self.rdoc.boxes[p].append(r)
        self.refresh_overlays()
        set_status(self, "Undone.")

    def clear_page(self):
        if not self.require_doc():
            return
        items = [(self.page_index, r) for r in self.rdoc.boxes[self.page_index]]
        if not items:
            set_status(self, "This page has no redactions.")
            return
        self.rdoc.boxes[self.page_index] = []
        self.undo_stack.append(("remove", items))
        self.refresh_overlays()
        set_status(self, f"🧹 Removed {len(items)} box(es) from this page (Undo brings them back).")

    # ---- search ----
    def do_search(self):
        if not self.require_doc():
            return
        query = self.search_edit.text().strip()
        if not query:
            self.search_hits = {}
            self.refresh_overlays()
            return
        self.search_hits, skipped = self.rdoc.find(query)
        count = sum(len(v) for v in self.search_hits.values())
        extra = f" ({skipped} scanned page(s) not searched - run Analyze first)" if skipped else ""
        if count:
            first = min(self.search_hits)
            if self.page_index not in self.search_hits:
                self.go_to(first)
            else:
                self.refresh_overlays()
            set_status(self, f"🔍 Found {count} match(es) on {len(self.search_hits)} page(s)." + extra)
        else:
            self.refresh_overlays()
            set_status(self, f"No matches for '{query}'." + extra)

    def redact_search_hits(self):
        if not self.require_doc():
            return
        if not self.search_hits:
            QMessageBox.information(self, APP_NAME, "Use 🔍 Find first, then ⬛ Redact All.")
            return
        items = [(p, r) for p, rects in self.search_hits.items() for r in rects]
        self.search_hits = {}
        self.add_boxes(items, "⬛ Redacted {n} match(es).")

    # ---- smart redaction ----
    def analyze(self):
        if not self.require_doc():
            return
        self.analyze_btn.setEnabled(False)
        self.analyze_progress.setVisible(True)
        self.analyze_progress.setRange(0, 0)
        self.found_label.setText("Analyzing…")
        self.worker = TaskWorker(analyze_for_redaction, self.rdoc.path, HAS_OCR)
        self.worker.progress.connect(self.on_analyze_progress)
        self.worker.done.connect(self.on_analyzed)
        self.worker.failed.connect(self.on_analyze_failed)
        self.worker.start()

    def on_analyze_progress(self, done, total, message):
        self.analyze_progress.setRange(0, max(total, 1))
        self.analyze_progress.setValue(done)
        set_status(self, message)

    def on_analyze_failed(self, message):
        self.analyze_btn.setEnabled(True)
        self.analyze_progress.setVisible(False)
        show_error(self, f"Analysis failed:\n\n{message}")

    def on_analyzed(self, result):
        self.analyze_btn.setEnabled(True)
        self.analyze_progress.setVisible(False)
        if result is None:
            return
        self.rdoc.ocr.update(result["ocr"])
        self.findings = [f for f in self.findings if f["cat"] == "Custom word"]
        self.findings = result["findings"] + self.findings
        self.fill_table()
        cats = len({f["cat"] for f in result["findings"]})
        self.found_label.setText(f"Found {len(result['findings'])} item(s) in {cats} categories")
        set_status(self, "🧠 Analysis complete. Tick the items to hide, then ✅ Apply.")
        if not result["findings"]:
            QMessageBox.information(self, APP_NAME, "✅ No sensitive information was detected.")

    def fill_table(self):
        self.table.setRowCount(0)
        colors = ["#1565c0", "#6a1b9a", "#c62828", "#2e7d32", "#ef6c00", "#00838f", "#ad1457"]
        for f in self.findings:
            r = self.table.rowCount()
            self.table.insertRow(r)
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            check.setCheckState(Qt.Checked if f["on"] else Qt.Unchecked)
            self.table.setItem(r, 0, check)
            cat = QTableWidgetItem(f["cat"])
            cat.setForeground(QColor(colors[CATEGORY_ORDER.get(f["cat"], 0) % len(colors)]))
            font = cat.font()
            font.setBold(True)
            cat.setFont(font)
            self.table.setItem(r, 1, cat)
            self.table.setItem(r, 2, QTableWidgetItem(f["text"]))
            pages = sorted({p + 1 for p, _ in f["hits"]})
            self.table.setItem(r, 3, QTableWidgetItem(describe_pages(pages)))
            self.table.setItem(r, 4, QTableWidgetItem(str(len(f["hits"]))))
            if f.get("applied"):
                for c in range(5):
                    self.table.item(r, c).setBackground(QColor("#e8f5e9"))

    def set_all_checks(self, on):
        for r in range(self.table.rowCount()):
            self.table.item(r, 0).setCheckState(Qt.Checked if on else Qt.Unchecked)

    def tick_selected_rows(self):
        rows = {i.row() for i in self.table.selectedItems()}
        for r in rows:
            self.table.item(r, 0).setCheckState(Qt.Checked)
        set_status(self, f"Ticked {len(rows)} highlighted row(s).")

    def on_finding_clicked(self, row, column):
        if column == 0 or row >= len(self.findings):
            return
        f = self.findings[row]
        hits = {}
        for p, r in f["hits"]:
            hits.setdefault(p, []).append(r)
        self.search_hits = hits
        self.go_to(f["hits"][0][0])
        set_status(self, f"Showing '{f['text']}' on page {f['hits'][0][0] + 1} "
                         "(double-click to open it in Manual Redaction).")

    def apply_selected(self):
        if not self.require_doc():
            return
        items, rows = [], 0
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0).checkState() == Qt.Checked and r < len(self.findings):
                f = self.findings[r]
                items += f["hits"]
                f["applied"], f["on"] = True, False
                rows += 1
        if not items:
            QMessageBox.information(self, APP_NAME, "Tick at least one item in the list first.")
            return
        self.search_hits = {}
        n = self.add_boxes(items, "✅ Applied {n} redaction box(es).")
        self.fill_table()
        QMessageBox.information(self, APP_NAME, f"✅ {rows} item(s) hidden with {n} box(es).\n\n"
                                                "Check the preview, then 💾 Save Redacted PDF.")

    def custom_words(self, add_to_list):
        if not self.require_doc():
            return
        words = [w.strip() for w in re.split(r"[,\n]", self.custom_edit.toPlainText())]
        words = list(dict.fromkeys(w for w in words if w))
        if not words:
            QMessageBox.information(self, APP_NAME, "Type one or more words first.")
            return
        new_findings, items, missing, skipped = [], [], [], 0
        for word in words:
            hits, skipped = self.rdoc.find(word)
            pairs = [(p, r) for p, rects in hits.items() for r in rects]
            if not pairs:
                missing.append(word)
                continue
            items += pairs
            new_findings.append({"cat": "Custom word", "text": word, "on": True, "hits": pairs})
        if missing:
            set_status(self, "Not found: " + ", ".join(missing[:8]))
        if add_to_list:
            known = {(f["cat"], f["text"].lower()) for f in self.findings}
            self.findings += [f for f in new_findings
                              if (f["cat"], f["text"].lower()) not in known]
            self.fill_table()
            msg = f"Added {len(new_findings)} word(s) to the list."
        else:
            n = self.add_boxes(items, "⬛ Redacted {n} custom word match(es).")
            msg = f"⬛ {n} match(es) of {len(new_findings)} word(s) redacted."
        if missing:
            msg += "\n\nNot found: " + ", ".join(missing)
        if skipped and not self.rdoc.ocr:
            msg += f"\n\n{skipped} scanned page(s) were not searched - run 🧠 Analyze first."
        QMessageBox.information(self, APP_NAME, msg)

    # ---- save ----
    def save_pdf(self):
        if not self.require_doc():
            return
        if self.rdoc.total_boxes() == 0:
            QMessageBox.information(self, APP_NAME, "There are no redactions to save yet.")
            return
        base = os.path.splitext(os.path.basename(self.rdoc.path))[0]
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save redacted PDF as",
            os.path.join(os.path.dirname(self.rdoc.path), f"{base}_redacted.pdf"),
            "PDF files (*.pdf)")
        if not out_path:
            return
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        if same_file(out_path, self.rdoc.path):
            show_error(self, "Please save under a new name - the original is kept unchanged.")
            return
        try:
            self.rdoc.save(out_path)
        except Exception as err:
            show_error(self, f"Saving failed:\n\n{err}")
            return
        self.unsaved = False
        set_status(self, f"💾 Saved {os.path.basename(out_path)}")
        show_success(self, f"✅ {self.rdoc.total_boxes()} redaction(s) applied permanently.\n\n"
                           "The text and images under the black boxes were removed from the "
                           f"file, not just covered.\n\nSaved to:\n{out_path}",
                     os.path.dirname(out_path))

    def shutdown(self):
        if self.worker is not None and self.worker.isRunning():
            self.worker.cancelled = True
            self.worker.wait(15000)


def scrollable(page):
    """Put a page inside a scroll area: on small screens it scrolls
    instead of squashing the buttons."""
    page.setObjectName("Page")
    page.setAttribute(Qt.WA_StyledBackground, True)
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.NoFrame)
    area.setWidget(page)
    return area


# ======================================================================
#  Main window
# ======================================================================

APP_STYLE = """
QMainWindow, #Page { background: #f3f4fb; }
QScrollArea { background: #f3f4fb; border: none; }
QScrollArea > QWidget > QWidget#Page { background: #f3f4fb; }
#Sidebar {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #2b1055, stop:0.6 #5b3fa8, stop:1 #7597de);
}
#AppTitle { color: white; font-size: 21px; font-weight: 800; background: transparent; }
#AppSub   { color: #e3dcff; font-size: 12px; background: transparent; }
#Version  { color: #d6ddff; font-size: 11px; background: transparent; }
QPushButton#NavButton {
    color: white; background: transparent; border: none; border-radius: 12px;
    text-align: left; padding: 10px 14px; font-size: 14px; font-weight: 600;
}
QPushButton#NavButton:hover   { background: rgba(255, 255, 255, 0.15); }
QPushButton#NavButton:checked { background: rgba(255, 255, 255, 0.30); }
#Card  { background: white; border-radius: 18px; }
#Step  { color: #2b1055; font-size: 15px; font-weight: 700; }
#Hint  { color: #6b7090; font-size: 13px; }
#Info  { color: #33395a; font-size: 14px; font-weight: 600; }
QLineEdit {
    border: 2px solid #d9def0; border-radius: 10px; padding: 10px 12px;
    font-size: 15px; background: #fbfbff; color: #222;
}
QLineEdit:focus { border-color: #8e54e9; background: white; }
QListWidget {
    border: 2px dashed #f2a3c4; border-radius: 12px; background: #fff8fb;
    padding: 6px; font-size: 14px; color: #222; outline: none;
}
QListWidget::item { padding: 9px; border-radius: 8px; margin: 2px; }
QListWidget::item:hover    { background: #ffe3ef; }
QListWidget::item:selected { background: #dd2476; color: white; }
QProgressBar {
    border: none; border-radius: 8px; background: #ececf6; height: 16px;
    text-align: center; color: #333;
}
QProgressBar::chunk {
    border-radius: 8px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff512f, stop:1 #dd2476);
}
QStatusBar { background: white; color: #5b3fa8; font-weight: 600; }
QMessageBox { background: white; }
QMessageBox QLabel { color: #222; font-size: 13px; }
QToolTip { color: #222; background: #fffbe6; border: 1px solid #e0c97f; }
QListWidget#PageGrid {
    border: 2px dashed #ffb366; background: #fffaf4; padding: 8px;
}
QListWidget#PageGrid::item { padding: 0px; margin: 0px; border-radius: 14px; }
QListWidget#PageGrid::item:hover    { background: #ffe9c7; }
QListWidget#PageGrid::item:selected { background: #ffd54f; }
#ToolPanel { background: #fff6ec; border-radius: 14px; }
QPlainTextEdit {
    border: 2px solid #f3c1d3; border-radius: 12px; padding: 8px;
    background: #fffafc; color: #222; font-family: Consolas, monospace; font-size: 13px;
}
QPlainTextEdit:focus { border-color: #b24592; }
QCheckBox { color: #33395a; font-size: 13px; spacing: 8px; }
QTabWidget::pane { border: 2px solid #bbdefb; border-radius: 12px; background: white; top: -2px; }
QTabBar::tab {
    background: #e8eaf6; color: #3f4470; padding: 8px 20px; margin-right: 4px;
    border-top-left-radius: 10px; border-top-right-radius: 10px;
    font-size: 13px; font-weight: 700;
}
QTabBar::tab:selected { background: white; color: #1565c0; border-bottom: 3px solid #2196f3; }
QTableWidget {
    border: 2px solid #e1e4f2; border-radius: 10px; background: white; color: #222;
    gridline-color: #eef0f7; font-size: 13px; selection-background-color: #bbdefb;
    selection-color: #111;
}
QHeaderView::section {
    background: #ede7f6; color: #311b92; font-weight: 700; padding: 6px;
    border: none; border-right: 1px solid #d1c4e9;
}
QTextBrowser {
    border: 2px solid #b2dfdb; border-radius: 12px; background: white; color: #222;
    padding: 6px; font-size: 13px;
}
QSlider::groove:horizontal { height: 6px; background: #ffe0b2; border-radius: 3px; }
QSlider::handle:horizontal {
    background: #ff9800; width: 16px; margin: -6px 0; border-radius: 8px;
}
#PreviewCard { background: white; border: 2px solid #c8e6c9; border-radius: 14px; }
#PreviewTitle { color: #2e7d32; font-size: 14px; font-weight: 800; }
#CanvasArea { background: #eceff8; border: 2px solid #bbdefb; border-radius: 10px; }
QSpinBox {
    border: 2px solid #ffd2a6; border-radius: 8px; padding: 5px;
    font-size: 14px; color: #222; background: white;
}
"""


class MainWindow(QMainWindow):
    def __init__(self, guard=None):
        super().__init__()
        self.guard = guard or UsageGuard(enforce=False)
        self.force_quit = False
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(900, 560)
        self.setStyleSheet(APP_STYLE)

        central = QWidget()
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- sidebar ----
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(230)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(14, 20, 14, 14)
        side.setSpacing(4)

        title = QLabel("PDF Studio Pro")
        title.setObjectName("AppTitle")
        subtitle = QLabel("Your all-in-one PDF toolkit")
        subtitle.setObjectName("AppSub")
        side.addWidget(title)
        side.addWidget(subtitle)
        side.addSpacing(14)

        self.ocr_page = OcrPage()
        self.redact_page = RedactionPage()
        self.pages = QStackedWidget()
        for page in (MergePage(), SplitPage(), ExtractPage(), OrganizePage(),
                     self.ocr_page, SummaryPage(), InvoicePage(), self.redact_page):
            self.pages.addWidget(scrollable(page))

        group = QButtonGroup(self)
        group.setExclusive(True)
        nav_items = ["🔗  Merge PDFs", "✂️  Split PDF", "📑  Extract Pages",
                     "🗂  Organize && Split", "🔤  OCR (Read Text)", "📝  1-Page Summary",
                     "🧾  Invoices to Excel", "⬛  Redaction"]
        for index, text in enumerate(nav_items):
            btn = QPushButton(text)
            btn.setObjectName("NavButton")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, i=index: self.pages.setCurrentIndex(i))
            group.addButton(btn)
            side.addWidget(btn)
            if index == 0:
                btn.setChecked(True)

        side.addStretch()
        self.trial_label = QLabel(self.guard.summary())
        self.trial_label.setObjectName("Version")
        self.trial_label.setVisible(self.guard.enforce)
        side.addWidget(self.trial_label)
        version = QLabel(f"Version {APP_VERSION}")
        version.setObjectName("Version")
        side.addWidget(version)

        root.addWidget(sidebar)
        root.addWidget(self.pages, 1)
        self.setCentralWidget(central)
        self.statusBar().showMessage("Ready. Choose a tool on the left to begin.")
        self.fit_to_screen()

        if self.guard.enforce:
            self.usage_timer = QTimer(self)
            self.usage_timer.timeout.connect(self.usage_tick)
            self.usage_timer.start(15000)        # count usage every 15 seconds
            if self.guard.days_left() <= 3:
                QTimer.singleShot(800, lambda: QMessageBox.information(
                    self, APP_NAME, f"Your trial ends in {self.guard.days_left()} day(s)."))

    def usage_tick(self):
        stop, warning = self.guard.tick()
        self.trial_label.setText(self.guard.summary())
        if stop:
            self.force_shutdown(stop)
        elif warning:
            box = QMessageBox(QMessageBox.Warning, APP_NAME, warning, QMessageBox.Ok, self)
            box.setModal(False)
            box.show()
            self.statusBar().showMessage(warning.splitlines()[0], 60000)

    def force_shutdown(self, message):
        self.usage_timer.stop()
        self.force_quit = True
        box = QMessageBox(QMessageBox.Warning, APP_NAME, message + "\n\nThe application "
                          "will now close.", QMessageBox.Ok, self)
        QTimer.singleShot(30000, box.accept)     # closes even if nobody clicks OK
        box.exec_()
        self.close()
        QApplication.quit()

    def fit_to_screen(self):
        """Open at a comfortable size that always fits the screen."""
        screen = QApplication.primaryScreen().availableGeometry()
        width = min(1200, screen.width() - 40)
        height = min(820, screen.height() - 60)
        self.resize(width, height)
        self.move(screen.x() + (screen.width() - width) // 2,
                  screen.y() + (screen.height() - height) // 2)

    def closeEvent(self, event):
        if not self.force_quit and self.redact_page.unsaved and QMessageBox.question(
                self, APP_NAME, "You have redactions that are not saved.\nClose anyway?") \
                != QMessageBox.Yes:
            event.ignore()
            return
        self.ocr_page.shutdown()
        for worker in list(TaskWorker.running):
            worker.cancelled = True
            worker.wait(15000)
        if self.guard.enforce and not self.force_quit:
            self.guard.tick()              # record the last minutes of use
        super().closeEvent(event)


def main():
    # In a windowed .exe there is no console: give libraries that print
    # progress (like the OCR engine) somewhere harmless to write.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")

    # Sharp text on high-resolution screens, using the exact Windows scale
    # (e.g. 150%) instead of rounding it up to 200%, which made it too big
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except AttributeError:
        pass  # very old PyQt5
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # Re-use an existing app object so re-running from IDLE (F5) is safe
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))
    app.setApplicationName(APP_NAME)

    guard = UsageGuard()
    blocked = guard.start_check()
    if blocked:
        QMessageBox.critical(None, APP_NAME, blocked)
        return

    window = MainWindow(guard)
    window.showMaximized()   # use the whole screen so no buttons are hidden
    app.exec_()   # no sys.exit() so IDLE does not ask to kill the shell


if __name__ == "__main__":
    main()
