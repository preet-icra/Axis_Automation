"""Add automatic internal links to index entries on PDF pages 2 and 3.

Install:
    python -m pip install --upgrade pymupdf

Run:
    python pdf_index_linker_pages_2_3.py
"""

from __future__ import annotations

import logging
import os
import re
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF

APP_TITLE = "PDF Index Linker - Pages 2 and 3"
INDEX_PAGE_NUMBERS = (1,2)  # Human-readable PDF page numbers
LOG_FILE = "pdf_index_linker.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


@dataclass(frozen=True)
class IndexEntry:
    source_page_index: int
    title: str
    printed_reference: str
    destination_page_number: int
    rectangle: fitz.Rect


def clean_title(text: str) -> str:
    """Remove dot leaders and trailing separators while retaining the entry name."""
    text = re.sub(r"[.\u2026\u2024\u00b7]{2,}.*$", "", text).strip()
    return text.strip(" .\u2026-:")


def cluster_lines_by_baseline(page: fitz.Page, tolerance: float = 1.0):
    """Combine text fragments that visually occupy the same index row."""
    fragments = []
    page_dict = page.get_text("dict")
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            text = " ".join(span.get("text", "") for span in line.get("spans", [])).strip()
            if text:
                fragments.append((fitz.Rect(line["bbox"]), text))

    rows: list[list[tuple[fitz.Rect, str]]] = []
    for rect, text in sorted(fragments, key=lambda item: (item[0].y0, item[0].x0)):
        matching_row = None
        for row in rows:
            baseline = sum(item[0].y0 for item in row) / len(row)
            if abs(rect.y0 - baseline) <= tolerance:
                matching_row = row
                break
        if matching_row is None:
            rows.append([(rect, text)])
        else:
            matching_row.append((rect, text))

    for row in rows:
        yield sorted(row, key=lambda item: item[0].x0)


def detect_index_entries(
    document: fitz.Document,
    index_page_numbers: tuple[int, ...] = INDEX_PAGE_NUMBERS,
) -> list[IndexEntry]:
    """Detect numbered index rows and use the first number of a range as target."""
    detected: list[IndexEntry] = []

    for page_number in index_page_numbers:
        page_index = page_number - 1
        if not 0 <= page_index < document.page_count:
            raise ValueError(
                f"Index page {page_number} does not exist. PDF has {document.page_count} pages."
            )

        page = document[page_index]
        for row in cluster_lines_by_baseline(page):
            combined_text = " ".join(text for _, text in row).strip()

            # Supports a single page, e.g. 56, and a range, e.g. 124 - 129.
            match = re.search(r"(?<!\d)(\d{1,3})(?:\s*[-\u2013]\s*(\d{1,3}))?\s*$", combined_text)
            if not match:
                continue

            destination = int(match.group(1))
            reference = match.group(0).strip()
            if not 1 <= destination <= document.page_count:
                continue

            number_fragment_index = None
            for position in range(len(row) - 1, -1, -1):
                if re.search(r"\d", row[position][1]):
                    number_fragment_index = position
                    break
            if number_fragment_index is None:
                continue

            # Factsheet index page numbers are right-aligned. This rejects years or
            # numbers that might occur inside ordinary body text.
            number_rect = row[number_fragment_index][0]
            if number_rect.x1 < page.rect.width * 0.80:
                continue

            title_text = re.sub(
                r"(?<!\d)\d{1,3}(?:\s*[-\u2013]\s*\d{1,3})?\s*$",
                "",
                combined_text,
            )
            title = clean_title(title_text)
            if not title or title.upper() == "INDEX":
                continue

            rectangle = fitz.Rect(row[0][0])
            for fragment_rect, _ in row[1:]:
                rectangle |= fragment_rect
            rectangle.x0 = max(0, rectangle.x0 - 2)
            rectangle.y0 = max(0, rectangle.y0 - 2)
            rectangle.x1 = min(page.rect.width, rectangle.x1 + 2)
            rectangle.y1 = min(page.rect.height, rectangle.y1 + 2)

            detected.append(
                IndexEntry(
                    source_page_index=page_index,
                    title=title,
                    printed_reference=reference,
                    destination_page_number=destination,
                    rectangle=rectangle,
                )
            )

    return detected


def add_index_links(source: Path, output: Path) -> tuple[int, dict[int, int]]:
    if not source.is_file():
        raise FileNotFoundError(f"Source PDF not found: {source}")
    if source.suffix.lower() != ".pdf":
        raise ValueError("The source file must be a PDF.")
    if source.resolve() == output.resolve():
        raise ValueError("Choose a different output filename to protect the original PDF.")

    document = fitz.open(source)
    try:
        if document.needs_pass:
            raise ValueError("The PDF is password protected and cannot be edited.")
        if document.page_count < max(INDEX_PAGE_NUMBERS):
            raise ValueError("The PDF does not contain both index pages 2 and 3.")

        entries = detect_index_entries(document)
        if not entries:
            raise RuntimeError("No right-aligned numbered index entries were detected.")

        per_page = {page_number: 0 for page_number in INDEX_PAGE_NUMBERS}
        for entry in entries:
            page = document[entry.source_page_index]
            page.insert_link(
                {
                    "kind": fitz.LINK_GOTO,
                    "from": entry.rectangle,
                    "page": entry.destination_page_number - 1,
                    "to": fitz.Point(0, 0),
                    "zoom": 0,
                }
            )
            per_page[entry.source_page_index + 1] += 1
            logging.info(
                "Index page %d: %s -> %s",
                entry.source_page_index + 1,
                entry.title,
                entry.printed_reference,
            )

        output.parent.mkdir(parents=True, exist_ok=True)
        temporary = output.with_name(f".{output.stem}.temporary.pdf")
        if temporary.exists():
            temporary.unlink()
        document.save(temporary, garbage=0, deflate=True)
    finally:
        document.close()

    try:
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink(missing_ok=True)

    # Reopen and validate the output structure and links.
    check = fitz.open(output)
    try:
        internal_links = 0
        validated_per_page = {}
        for page_number in INDEX_PAGE_NUMBERS:
            links = [
                link
                for link in check[page_number - 1].get_links()
                if link.get("kind") == fitz.LINK_GOTO
                and 0 <= link.get("page", -1) < check.page_count
            ]
            validated_per_page[page_number] = len(links)
            internal_links += len(links)

        if internal_links < len(entries):
            raise RuntimeError(
                f"Validation failed: expected at least {len(entries)} links, "
                f"but found {internal_links}."
            )
    finally:
        check.close()

    return len(entries), per_page


class IndexLinkerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x330")
        self.minsize(680, 300)
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select the source PDF.")
        self.build_ui()

    def build_ui(self) -> None:
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Source PDF").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.source_var).grid(
            row=0, column=1, sticky="ew", padx=10, pady=8
        )
        ttk.Button(frame, text="Browse...", command=self.choose_source).grid(row=0, column=2)

        ttk.Label(frame, text="Output PDF").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.output_var).grid(
            row=1, column=1, sticky="ew", padx=10, pady=8
        )
        ttk.Button(frame, text="Save as...", command=self.choose_output).grid(row=1, column=2)

        explanation = (
            "The program automatically scans actual PDF pages 2 and 3. Each numbered "
            "index row becomes clickable and opens the first page number printed on that row. "
            "For a range such as 124-129, it opens page 124."
        )
        ttk.Label(frame, text=explanation, wraplength=700).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(12, 8)
        )

        ttk.Separator(frame).grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
        ttk.Label(frame, textvariable=self.status_var).grid(
            row=4, column=0, columnspan=2, sticky="w"
        )
        ttk.Button(frame, text="Create linked PDF", command=self.process).grid(
            row=4, column=2, sticky="e"
        )

    def choose_source(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select source PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.source_var.set(selected)
        source = Path(selected)
        self.output_var.set(str(source.with_name(f"{source.stem}_linked_index.pdf")))
        self.status_var.set("Source selected. Choose output location or create the PDF.")

    def choose_output(self) -> None:
        current = self.output_var.get().strip()
        selected = filedialog.asksaveasfilename(
            title="Save linked PDF as",
            defaultextension=".pdf",
            initialfile=Path(current).name if current else "linked_index.pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if selected:
            self.output_var.set(selected)

    def process(self) -> None:
        try:
            source_text = self.source_var.get().strip()
            output_text = self.output_var.get().strip()
            if not source_text:
                raise ValueError("Select a source PDF.")
            if not output_text:
                raise ValueError("Select an output PDF location.")

            source = Path(source_text).expanduser()
            output = Path(output_text).expanduser()
            if output.suffix.lower() != ".pdf":
                output = output.with_suffix(".pdf")
                self.output_var.set(str(output))

            self.status_var.set("Processing index pages 2 and 3...")
            self.update_idletasks()
            total, per_page = add_index_links(source, output)
            summary = ", ".join(
                f"page {page}: {count} links" for page, count in per_page.items()
            )
            self.status_var.set(f"Created {total} links ({summary}).")
            messagebox.showinfo(
                "Success",
                f"Linked PDF created:\n{output}\n\n{total} links added ({summary}).",
            )
        except Exception as exc:
            logging.exception("PDF processing failed")
            self.status_var.set("Processing failed. See the message for details.")
            messagebox.showerror("Could not create linked PDF", str(exc))


if __name__ == "__main__":
    IndexLinkerApp().mainloop()
