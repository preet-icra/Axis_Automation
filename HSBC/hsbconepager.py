"""
HSBC Fund-wise PDF Splitter and Disclaimer Merger

Windows installation:
    py -m pip install PyQt6 PyMuPDF pypdf pywin32

Run:
    py hsbc_fund_pdf_tool.py

The application has two tabs:
1. Split by Heading: detects fund headings such as "HSBC Large Cap Fund",
   groups consecutive pages under the detected fund, and creates one PDF per fund.
2. Merge Disclaimer: appends or prepends a disclaimer PDF to every fund PDF.

Notes:
- Text-based PDFs work directly.
- For scanned/image-only PDFs, install the Tesseract OCR desktop program and enable OCR.
  Python OCR packages: py -m pip install pytesseract Pillow
  The app auto-detects common Windows locations and TESSERACT_CMD.
"""

from __future__ import annotations

import os
import re
import sys
import traceback
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PyQt6.QtCore import QObject, QThread, pyqtSignal
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtCore import QUrl
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

try:
    import win32api  # pywin32, used for native Windows launch/open behavior
except ImportError:
    win32api = None


APP_TITLE = "Red Heading PDF Splitter and Disclaimer Merger v7"
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]')

# The first expression covers the layout visible in the supplied image:
# "HSBC Large Cap Fund". Additional expressions support common AMC wording.
DEFAULT_PATTERNS = [
    r"\b(HSBC\s+[A-Z][A-Za-z0-9&'().,+/\- ]{2,90}?\s+Fund)\b",
    r"\b([A-Z][A-Za-z0-9&'().,+/\- ]{2,90}?\s+(?:Mutual\s+)?Fund)\b",
    r"(?:Scheme|Fund)\s+Name\s*[:\-]\s*([^\r\n]{3,100})",
]


@dataclass
class FundRange:
    name: str
    start_page: int  # zero-based, inclusive
    end_page: int    # zero-based, inclusive


class CancelledError(Exception):
    pass


def clean_fund_name(raw: str) -> str:
    """Normalize extracted heading and make it suitable for a Windows filename."""
    text = re.sub(r"\s+", " ", raw or "").strip(" .-_\t\r\n")
    # Remove text occasionally captured after the heading on the same extracted line.
    text = re.split(
        r"\s{2,}|(?=\b(?:Investment Objective|Fund Details|Benchmark|Date of Allotment|NAV)\b)",
        text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()
    text = INVALID_FILENAME_CHARS.sub("-", text)
    text = re.sub(r"\s*-\s*", " - ", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:100] or "Unknown Fund"


def unique_path(folder: Path, stem: str, suffix: str = ".pdf") -> Path:
    candidate = folder / f"{stem}{suffix}"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{stem} ({counter}){suffix}"
        counter += 1
    return candidate


def find_tesseract_executable() -> Optional[Path]:
    """Find tesseract.exe from PATH, environment variables, or common Windows folders."""
    import shutil

    candidates: list[Path] = []
    path_match = shutil.which("tesseract")
    if path_match:
        candidates.append(Path(path_match))

    for variable in ("TESSERACT_CMD", "TESSERACT_PATH"):
        value = os.environ.get(variable, "").strip().strip('\"')
        if value:
            candidate = Path(value)
            candidates.append(candidate / "tesseract.exe" if candidate.is_dir() else candidate)

    for root_variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(root_variable)
        if root:
            candidates.extend([
                Path(root) / "Tesseract-OCR" / "tesseract.exe",
                Path(root) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
            ])

    candidates.extend([
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ])

    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def ocr_page_text(page: fitz.Page, dpi: int = 200) -> str:
    """OCR fallback for image-only pages. Requires Pillow, pytesseract and Tesseract OCR."""
    try:
        import io
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError(
            "OCR Python packages are missing. Run: py -m pip install pytesseract Pillow"
        ) from exc

    executable = find_tesseract_executable()
    if executable is None:
        raise RuntimeError(
            "The pytesseract Python package is installed, but the Tesseract OCR program "
            "(tesseract.exe) is not installed or cannot be found. Install Tesseract OCR, "
            "then restart the application. The usual path is "
            "C:\\Program Files\\Tesseract-OCR\\tesseract.exe. Alternatively, set the "
            "TESSERACT_CMD environment variable to the full tesseract.exe path. If the PDF "
            "has selectable text, untick 'Use OCR' and run again."
        )

    pytesseract.pytesseract.tesseract_cmd = str(executable)
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    image = Image.open(io.BytesIO(pix.tobytes("png")))
    try:
        return pytesseract.image_to_string(image)
    except pytesseract.TesseractNotFoundError as exc:
        raise RuntimeError(
            f"Tesseract OCR could not start from: {executable}. Reinstall Tesseract OCR "
            "or set TESSERACT_CMD to the correct tesseract.exe file."
        ) from exc


def detect_fund_name(text: str, custom_regex: str = "") -> Optional[str]:
    if not text or not text.strip():
        return None

    patterns = [custom_regex] if custom_regex.strip() else DEFAULT_PATTERNS
    candidates: list[tuple[int, str]] = []

    for pattern_index, pattern in enumerate(patterns):
        try:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE | re.MULTILINE):
                value = match.group(1) if match.lastindex else match.group(0)
                value = clean_fund_name(value)
                # Ignore the AMC/logo text, which often appears before the actual scheme title.
                if value.casefold() in {"hsbc mutual fund", "mutual fund"}:
                    continue
                if len(value) < 4:
                    continue
                # Lower score is better. Prefer HSBC and earlier, specific patterns.
                score = pattern_index * 10000 + match.start()
                if value.lower().startswith("hsbc "):
                    score -= 5000
                candidates.append((score, value))
        except re.error as exc:
            raise ValueError(f"Invalid custom regular expression: {exc}") from exc

    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def pdf_color_to_rgb(color_value: int) -> tuple[int, int, int]:
    """Convert PyMuPDF span color (0xRRGGBB) to an RGB tuple."""
    value = int(color_value or 0)
    return (value >> 16) & 255, (value >> 8) & 255, value & 255


def is_red_color(color_value: int) -> bool:
    """Accept bright/dark red text while rejecting black, grey, orange and blue."""
    red, green, blue = pdf_color_to_rgb(color_value)
    return (
        red >= 125
        and red >= green + 45
        and red >= blue + 45
        and red >= green * 1.45
        and red >= blue * 1.45
    )


def detect_heading_from_page(page: fitz.Page, custom_regex: str = "") -> Optional[str]:
    """Detect the principal large red heading in the upper half of a page."""
    page_text = page.get_text("text", sort=True)
    if custom_regex.strip():
        return detect_fund_name(page_text, custom_regex)

    page_height = float(page.rect.height or 1)
    data = page.get_text("dict", sort=True)
    red_spans: list[dict] = []

    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = re.sub(r"\s+", " ", span.get("text", "")).strip()
                bbox = span.get("bbox", [0, 0, 0, 0])
                size = float(span.get("size", 0.0))
                color = int(span.get("color", 0))
                y0 = float(bbox[1])
                if (
                    text
                    and y0 <= page_height * 0.50
                    and size >= 11.0
                    and is_red_color(color)
                ):
                    red_spans.append({
                        "text": text,
                        "x0": float(bbox[0]),
                        "y0": y0,
                        "y1": float(bbox[3]),
                        "size": size,
                    })

    if not red_spans:
        return None

    # Join red spans on the same baseline into visual rows.
    red_spans.sort(key=lambda item: (item["y0"], item["x0"]))
    rows: list[dict] = []
    for span in red_spans:
        tolerance = max(3.0, span["size"] * 0.38)
        row = next(
            (item for item in rows if abs(item["y0"] - span["y0"]) <= tolerance),
            None,
        )
        if row is None:
            row = {
                "y0": span["y0"],
                "y1": span["y1"],
                "max_size": span["size"],
                "parts": [],
            }
            rows.append(row)
        row["parts"].append((span["x0"], span["text"]))
        row["max_size"] = max(row["max_size"], span["size"])
        row["y1"] = max(row["y1"], span["y1"])

    for row in rows:
        row["text"] = re.sub(
            r"\s+", " ", " ".join(text for _, text in sorted(row["parts"]))
        ).strip(" .:")

    # Principal heading is normally the largest red row. Prefer HSBC on a tie.
    valid_rows = [
        row for row in rows
        if 4 <= len(row["text"]) <= 160
        and row["text"].casefold() not in {"hsbc mutual fund", "mutual fund"}
    ]
    if not valid_rows:
        return None

    valid_rows.sort(
        key=lambda row: (
            -row["max_size"],
            0 if row["text"].casefold().startswith("hsbc ") else 1,
            row["y0"],
        )
    )
    first = valid_rows[0]
    heading = first["text"]

    # Join a wrapped second red line when size and vertical spacing match.
    following = sorted(
        (row for row in valid_rows if row["y0"] > first["y0"]),
        key=lambda row: row["y0"],
    )
    if following:
        second = following[0]
        vertical_gap = second["y0"] - first["y1"]
        similar_size = second["max_size"] >= first["max_size"] * 0.75
        close = vertical_gap <= first["max_size"] * 1.10
        if similar_size and close and len(heading) + len(second["text"]) < 160:
            if heading.endswith("-"):
                heading = heading[:-1].rstrip() + " " + second["text"]
            else:
                heading = heading + " " + second["text"]

    return clean_fund_name(heading)

def normalize_output_folder(folder: Path) -> Path:
    """Validate a normal Windows folder without constructing a broken \\? path."""
    raw = os.path.expandvars(os.path.expanduser(str(folder))).strip().strip('"')
    if not raw or raw in {".", "\\", "/"}:
        raise RuntimeError("Please select a valid output folder.")
    normalized = Path(os.path.normpath(raw))
    try:
        normalized.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise RuntimeError(
            f"Windows cannot create or access this output folder:\n{normalized}\n\n"
            "Select a short local folder such as C:\\HSBC_Output\\Split.\n"
            f"Windows error: {exc}"
        ) from exc
    if not normalized.is_dir():
        raise RuntimeError(f"The selected output location is not a folder: {normalized}")
    return normalized


def windows_safe_output_path(folder: Path, stem: str) -> Path:
    folder = normalize_output_folder(folder)
    safe_stem = clean_fund_name(stem)
    max_total = 235 if os.name == "nt" else 4096
    available = max(24, max_total - len(str(folder)) - 16)
    safe_stem = safe_stem[:available].rstrip(" .") or "Heading"
    return unique_path(folder, safe_stem)


def write_page_range_with_pypdf(source_pdf: Path, destination: Path, start_page: int, end_page: int) -> None:
    """Create locally, then copy to the destination. No malformed extended path prefix."""
    import shutil
    import tempfile

    destination.parent.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(source_pdf))
    writer = PdfWriter()
    for page_number in range(start_page, end_page + 1):
        writer.add_page(reader.pages[page_number])

    fd, temp_name = tempfile.mkstemp(prefix="pdf_split_", suffix=".pdf")
    os.close(fd)
    try:
        with open(temp_name, "wb") as handle:
            writer.write(handle)
        if os.path.getsize(temp_name) <= 0:
            raise RuntimeError("The temporary PDF is empty.")
        shutil.copyfile(temp_name, str(destination))
        if not destination.is_file() or destination.stat().st_size <= 0:
            raise RuntimeError(f"Output PDF was not created: {destination}")
    except OSError as exc:
        raise RuntimeError(
            f"Could not save the PDF to:\n{destination}\n\n"
            "Use a short local output folder such as C:\\HSBC_Output\\Split. "
            "Do not select a PDF filename as the output folder.\n"
            f"Windows error: {exc}"
        ) from exc
    finally:
        try:
            os.remove(temp_name)
        except OSError:
            pass

def find_fund_ranges(
    source_pdf: Path,
    custom_regex: str,
    use_ocr: bool,
    ocr_dpi: int,
    progress: Callable[[int, str], None],
    is_cancelled: Callable[[], bool],
) -> tuple[list[FundRange], list[str]]:
    doc = fitz.open(source_pdf)
    logs: list[str] = []
    starts: list[tuple[int, str]] = []
    current_name: Optional[str] = None

    try:
        total = len(doc)
        for page_index, page in enumerate(doc):
            if is_cancelled():
                raise CancelledError()

            progress(int((page_index / max(total, 1)) * 55), f"Reading page {page_index + 1} of {total}")
            text = page.get_text("text", sort=True)
            ocr_was_used = False
            if use_ocr and len(text.strip()) < 25:
                logs.append(f"Page {page_index + 1}: little/no embedded text, running OCR.")
                text = ocr_page_text(page, ocr_dpi)
                ocr_was_used = True

            detected = detect_heading_from_page(page, custom_regex)
            if detected is None and ocr_was_used:
                # OCR loses font/layout metadata, so use a strict HSBC-only expression.
                detected = detect_fund_name(
                    text,
                    custom_regex or r"(?mi)^\s*(HSBC\s+[^\r\n]{2,100}?\s+Fund)\s*$",
                )
            if detected and detected.casefold() != (current_name or "").casefold():
                starts.append((page_index, detected))
                current_name = detected
                logs.append(f"Page {page_index + 1}: detected '{detected}'.")
            elif detected:
                logs.append(f"Page {page_index + 1}: continues '{current_name}'.")
            else:
                logs.append(f"Page {page_index + 1}: no new large red heading found.")

        if not starts:
            raise RuntimeError(
                "No large red heading was detected. Try a custom regular expression or enable OCR for scanned PDFs."
            )

        ranges: list[FundRange] = []
        # Pages before the first detected fund are intentionally not attached to a fund.
        if starts[0][0] > 0:
            logs.append(
                f"Pages 1-{starts[0][0]} occur before the first detected heading and will be skipped."
            )
        for index, (start_page, name) in enumerate(starts):
            end_page = starts[index + 1][0] - 1 if index + 1 < len(starts) else total - 1
            ranges.append(FundRange(name, start_page, end_page))
        return ranges, logs
    finally:
        doc.close()


def split_pdf_by_fund(
    source_pdf: Path,
    output_dir: Path,
    custom_regex: str,
    use_ocr: bool,
    ocr_dpi: int,
    progress: Callable[[int, str], None],
    is_cancelled: Callable[[], bool],
) -> list[Path]:
    output_dir = normalize_output_folder(output_dir)
    ranges, logs = find_fund_ranges(
        source_pdf, custom_regex, use_ocr, ocr_dpi, progress, is_cancelled
    )
    for line in logs:
        progress(-1, line)

    created: list[Path] = []
    used_names: dict[str, int] = {}
    total_ranges = len(ranges)
    for index, fund_range in enumerate(ranges):
        if is_cancelled():
            raise CancelledError()

        base_name = clean_fund_name(fund_range.name)
        key = base_name.casefold()
        used_names[key] = used_names.get(key, 0) + 1
        if used_names[key] > 1:
            base_name = f"{base_name} - Part {used_names[key]}"

        output_path = windows_safe_output_path(output_dir, base_name)
        write_page_range_with_pypdf(
            source_pdf, output_path, fund_range.start_page, fund_range.end_page
        )
        created.append(output_path)
        progress(
            55 + int(((index + 1) / total_ranges) * 45),
            f"Created {output_path.name} from pages "
            f"{fund_range.start_page + 1}-{fund_range.end_page + 1}",
        )
    return created


def merge_one_pdf(fund_pdf: Path, disclaimer_pdf: Path, destination: Path, position: str) -> None:
    writer = PdfWriter()
    ordered = (
        [disclaimer_pdf, fund_pdf]
        if position == "Before fund pages"
        else [fund_pdf, disclaimer_pdf]
    )
    for pdf_path in ordered:
        reader = PdfReader(str(pdf_path))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception as exc:
                raise RuntimeError(f"Cannot open encrypted PDF: {pdf_path.name}") from exc
        for page in reader.pages:
            writer.add_page(page)

    with destination.open("wb") as file_handle:
        writer.write(file_handle)


def merge_disclaimer_batch(
    input_dir: Path,
    disclaimer_pdf: Path,
    output_dir: Path,
    position: str,
    month_name: str,
    progress: Callable[[int, str], None],
    is_cancelled: Callable[[], bool],
) -> list[Path]:
    output_dir = normalize_output_folder(output_dir)
    fund_files = sorted(
        p for p in input_dir.glob("*.pdf")
        if p.resolve() != disclaimer_pdf.resolve()
    )
    if not fund_files:
        raise RuntimeError("No fund PDF files were found in the selected input folder.")

    selected_month = re.sub(r"\s+", " ", month_name).strip()
    if not selected_month:
        raise RuntimeError("Please enter or select a month name.")
    if INVALID_FILENAME_CHARS.search(selected_month):
        raise RuntimeError("The month name contains a character that is not allowed in Windows filenames.")
    month_year = f"{selected_month} {datetime.now().year}"

    created: list[Path] = []
    for index, fund_pdf in enumerate(fund_files):
        if is_cancelled():
            raise CancelledError()
        # Final naming format: Fund Name - Current Month Year.pdf
        # Example when run in August 2026: HSBC Large Cap Fund - August 2026.pdf
        fund_name = re.sub(
            r"\s+-\s+(?:With Disclaimer|(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})$",
            "",
            fund_pdf.stem,
            flags=re.IGNORECASE,
        ).strip()
        stem = f"{fund_name} - {month_year}"
        destination = windows_safe_output_path(output_dir, stem)
        merge_one_pdf(fund_pdf, disclaimer_pdf, destination, position)
        created.append(destination)
        progress(
            int(((index + 1) / len(fund_files)) * 100),
            f"Created {destination.name}",
        )
    return created


class JobWorker(QObject):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, function: Callable, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs
        self._cancelled = False

    def request_cancel(self):
        self._cancelled = True

    def is_cancelled(self) -> bool:
        return self._cancelled

    def report(self, value: int, message: str):
        self.progress.emit(value, message)

    def run(self):
        try:
            result = self.function(
                *self.args,
                progress=self.report,
                is_cancelled=self.is_cancelled,
                **self.kwargs,
            )
            if self._cancelled:
                self.cancelled.emit()
            else:
                self.finished.emit(result)
        except CancelledError:
            self.cancelled.emit()
        except Exception:
            self.failed.emit(traceback.format_exc())


class PathRow(QWidget):
    def __init__(self, mode: str, caption: str, file_filter: str = "PDF files (*.pdf)"):
        super().__init__()
        self.mode = mode
        self.caption = caption
        self.file_filter = file_filter
        self.edit = QLineEdit()
        self.button = QPushButton("Browse...")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.edit, 1)
        layout.addWidget(self.button)
        self.button.clicked.connect(self.browse)

    def browse(self):
        start = self.edit.text().strip() or str(Path.home())
        if self.mode == "file":
            value, _ = QFileDialog.getOpenFileName(self, self.caption, start, self.file_filter)
        else:
            value = QFileDialog.getExistingDirectory(self, self.caption, start)
        if value:
            self.edit.setText(value)

    def path(self) -> Path:
        return Path(self.edit.text().strip())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(900, 680)
        self.thread: Optional[QThread] = None
        self.worker: Optional[JobWorker] = None
        self.last_output_dir: Optional[Path] = None

        self.tabs = QTabWidget()
        self.tabs.addTab(self.build_split_tab(), "1. Split by Heading")
        self.tabs.addTab(self.build_merge_tab(), "2. Merge Disclaimer")
        self.setCentralWidget(self.tabs)
        self.statusBar().showMessage("Ready")

    def build_split_tab(self) -> QWidget:
        tab = QWidget()
        outer = QVBoxLayout(tab)

        intro = QLabel(
            "Select the complete scheme PDF. The program reads each page, detects headings "
            "such as 'HSBC Large Cap Fund', groups consecutive pages, and saves one PDF per fund."
        )
        intro.setWordWrap(True)
        outer.addWidget(intro)

        box = QGroupBox("Split settings")
        form = QFormLayout(box)
        self.split_source = PathRow("file", "Select complete PDF")
        self.split_output = PathRow("folder", "Select output folder")
        self.custom_regex = QLineEdit()
        self.custom_regex.setPlaceholderText(
            r"Optional, for example: \b(HSBC\s+.+?\s+Fund)\b"
        )
        self.use_ocr = QCheckBox("Use OCR when a page contains little or no embedded text")
        self.ocr_dpi = QSpinBox()
        self.ocr_dpi.setRange(120, 400)
        self.ocr_dpi.setValue(200)

        form.addRow("Complete source PDF:", self.split_source)
        form.addRow("Split PDF output folder:", self.split_output)
        form.addRow("Custom heading regex:", self.custom_regex)
        form.addRow("Scanned PDF support:", self.use_ocr)
        form.addRow("OCR resolution (DPI):", self.ocr_dpi)
        outer.addWidget(box)

        buttons = QHBoxLayout()
        self.split_run = QPushButton("Split PDF by Heading")
        self.split_cancel = QPushButton("Cancel")
        self.split_cancel.setEnabled(False)
        self.split_open = QPushButton("Open Output Folder")
        self.split_open.setEnabled(False)
        buttons.addWidget(self.split_run)
        buttons.addWidget(self.split_cancel)
        buttons.addStretch(1)
        buttons.addWidget(self.split_open)
        outer.addLayout(buttons)

        self.split_progress = QProgressBar()
        self.split_log = QPlainTextEdit()
        self.split_log.setReadOnly(True)
        outer.addWidget(self.split_progress)
        outer.addWidget(QLabel("Processing log:"))
        outer.addWidget(self.split_log, 1)

        self.split_run.clicked.connect(self.start_split)
        self.split_cancel.clicked.connect(self.cancel_job)
        self.split_open.clicked.connect(self.open_last_output)
        return tab

    def build_merge_tab(self) -> QWidget:
        tab = QWidget()
        outer = QVBoxLayout(tab)
        intro = QLabel(
            "Select the folder containing the split fund PDFs and one disclaimer PDF. "
            "The program creates a new PDF for every fund with the disclaimer added and names it "
            "as Fund Name - Selected Month Current Year.pdf."
        )
        intro.setWordWrap(True)
        outer.addWidget(intro)

        box = QGroupBox("Merge settings")
        form = QFormLayout(box)
        self.merge_input = PathRow("folder", "Select folder containing fund PDFs")
        self.disclaimer_file = PathRow("file", "Select disclaimer PDF")
        self.merge_output = PathRow("folder", "Select merged PDF output folder")
        self.disclaimer_position = QComboBox()
        self.disclaimer_position.addItems(["After fund pages", "Before fund pages"])
        self.merge_month = QComboBox()
        self.merge_month.setEditable(True)
        self.merge_month.addItems([
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December",
        ])
        self.merge_month.setCurrentText(datetime.now().strftime("%B"))
        self.merge_month.setToolTip("Select a month or type a custom month name.")

        form.addRow("Fund PDF input folder:", self.merge_input)
        form.addRow("Disclaimer PDF:", self.disclaimer_file)
        form.addRow("Merged PDF output folder:", self.merge_output)
        form.addRow("Disclaimer position:", self.disclaimer_position)
        form.addRow("Month name:", self.merge_month)
        self.naming_example = QLabel()
        form.addRow("Filename example:", self.naming_example)
        self.merge_month.currentTextChanged.connect(self.update_merge_filename_example)
        self.update_merge_filename_example()
        outer.addWidget(box)

        buttons = QHBoxLayout()
        self.merge_run = QPushButton("Merge Disclaimer into All Fund PDFs")
        self.merge_cancel = QPushButton("Cancel")
        self.merge_cancel.setEnabled(False)
        self.merge_open = QPushButton("Open Output Folder")
        self.merge_open.setEnabled(False)
        buttons.addWidget(self.merge_run)
        buttons.addWidget(self.merge_cancel)
        buttons.addStretch(1)
        buttons.addWidget(self.merge_open)
        outer.addLayout(buttons)

        self.merge_progress = QProgressBar()
        self.merge_log = QPlainTextEdit()
        self.merge_log.setReadOnly(True)
        outer.addWidget(self.merge_progress)
        outer.addWidget(QLabel("Processing log:"))
        outer.addWidget(self.merge_log, 1)

        self.merge_run.clicked.connect(self.start_merge)
        self.merge_cancel.clicked.connect(self.cancel_job)
        self.merge_open.clicked.connect(self.open_last_output)
        return tab

    def validate_file(self, path: Path, label: str) -> bool:
        if not path.is_file() or path.suffix.lower() != ".pdf":
            QMessageBox.warning(self, APP_TITLE, f"Please select a valid PDF for {label}.")
            return False
        return True

    def start_split(self):
        source = self.split_source.path()
        output = self.split_output.path()
        if not self.validate_file(source, "the complete source PDF"):
            return
        if not self.split_output.edit.text().strip():
            QMessageBox.warning(self, APP_TITLE, "Please select an output folder.")
            return
        try:
            if self.custom_regex.text().strip():
                re.compile(self.custom_regex.text().strip())
        except re.error as exc:
            QMessageBox.warning(self, APP_TITLE, f"Invalid custom regular expression:\n{exc}")
            return

        self.split_log.clear()
        self.split_progress.setValue(0)
        self.last_output_dir = output
        self.run_job(
            split_pdf_by_fund,
            (
                source,
                output,
                self.custom_regex.text().strip(),
                self.use_ocr.isChecked(),
                self.ocr_dpi.value(),
            ),
            self.split_progress,
            self.split_log,
            self.split_run,
            self.split_cancel,
            self.split_open,
        )

    def update_merge_filename_example(self):
        month = re.sub(r"\s+", " ", self.merge_month.currentText()).strip()
        month = month or "Month"
        self.naming_example.setText(
            f"HSBC Large Cap Fund - {month} {datetime.now().year}.pdf"
        )

    def start_merge(self):
        input_dir = self.merge_input.path()
        disclaimer = self.disclaimer_file.path()
        output = self.merge_output.path()
        if not input_dir.is_dir():
            QMessageBox.warning(self, APP_TITLE, "Please select a valid fund PDF input folder.")
            return
        if not self.validate_file(disclaimer, "the disclaimer"):
            return
        if not self.merge_output.edit.text().strip():
            QMessageBox.warning(self, APP_TITLE, "Please select an output folder.")
            return
        if output.resolve() == input_dir.resolve():
            QMessageBox.warning(
                self,
                APP_TITLE,
                "Use a separate output folder to avoid processing previously merged PDFs.",
            )
            return

        month_name = re.sub(r"\s+", " ", self.merge_month.currentText()).strip()
        if not month_name:
            QMessageBox.warning(self, APP_TITLE, "Please enter or select a month name.")
            return
        if INVALID_FILENAME_CHARS.search(month_name):
            QMessageBox.warning(
                self,
                APP_TITLE,
                "Month name contains a character that is not allowed in a Windows filename.",
            )
            return

        self.merge_log.clear()
        self.merge_progress.setValue(0)
        self.last_output_dir = output
        self.run_job(
            merge_disclaimer_batch,
            (
                input_dir,
                disclaimer,
                output,
                self.disclaimer_position.currentText(),
                month_name,
            ),
            self.merge_progress,
            self.merge_log,
            self.merge_run,
            self.merge_cancel,
            self.merge_open,
        )

    def run_job(self, function, args, progress_bar, log, run_button, cancel_button, open_button):
        if self.thread is not None:
            return
        run_button.setEnabled(False)
        cancel_button.setEnabled(True)
        open_button.setEnabled(False)
        self.statusBar().showMessage("Processing...")

        self.thread = QThread(self)
        self.worker = JobWorker(function, *args)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(
            lambda value, message: self.on_progress(progress_bar, log, value, message)
        )
        self.worker.finished.connect(
            lambda files: self.on_success(files, log, run_button, cancel_button, open_button)
        )
        self.worker.failed.connect(
            lambda details: self.on_failure(details, log, run_button, cancel_button)
        )
        self.worker.cancelled.connect(
            lambda: self.on_cancelled(log, run_button, cancel_button)
        )
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.worker.cancelled.connect(self.thread.quit)
        self.thread.finished.connect(self.cleanup_thread)
        self.thread.start()

    def on_progress(self, progress_bar, log, value: int, message: str):
        if value >= 0:
            progress_bar.setValue(max(0, min(100, value)))
        log.appendPlainText(message)
        self.statusBar().showMessage(message)

    def on_success(self, files, log, run_button, cancel_button, open_button):
        run_button.setEnabled(True)
        cancel_button.setEnabled(False)
        open_button.setEnabled(True)
        log.appendPlainText(f"Completed. Created {len(files)} PDF file(s).")
        self.statusBar().showMessage("Completed")
        QMessageBox.information(
            self,
            APP_TITLE,
            f"Completed successfully.\n\nCreated {len(files)} PDF file(s).",
        )

    def on_failure(self, details, log, run_button, cancel_button):
        run_button.setEnabled(True)
        cancel_button.setEnabled(False)
        log.appendPlainText(details)
        self.statusBar().showMessage("Failed")
        last_line = next((line for line in reversed(details.splitlines()) if line.strip()), "Unknown error")
        QMessageBox.critical(self, APP_TITLE, f"Processing failed.\n\n{last_line}")

    def on_cancelled(self, log, run_button, cancel_button):
        run_button.setEnabled(True)
        cancel_button.setEnabled(False)
        log.appendPlainText("Operation cancelled by user.")
        self.statusBar().showMessage("Cancelled")

    def cancel_job(self):
        if self.worker is not None:
            self.worker.request_cancel()
            self.statusBar().showMessage("Cancelling after the current page/file...")

    def cleanup_thread(self):
        if self.worker is not None:
            self.worker.deleteLater()
        if self.thread is not None:
            self.thread.deleteLater()
        self.worker = None
        self.thread = None

    def open_last_output(self):
        if not self.last_output_dir:
            return
        folder = self.last_output_dir
        folder.mkdir(parents=True, exist_ok=True)
        try:
            # Explicit pywin32 usage on Windows, as requested.
            if os.name == "nt" and win32api is not None:
                win32api.ShellExecute(0, "open", str(folder), None, None, 1)
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
        except Exception as exc:
            QMessageBox.warning(self, APP_TITLE, f"Could not open the folder:\n{exc}")

    def closeEvent(self, event):
        if self.worker is not None:
            self.worker.request_cancel()
            QMessageBox.information(
                self,
                APP_TITLE,
                "A job is still running. It will stop after the current page or file.",
            )
            event.ignore()
            return
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
