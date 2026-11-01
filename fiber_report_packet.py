"""Tkinter utility for preparing an abaca-fiber proforma invoice packet.

PDF-XChange Editor is used as an external preview application. ReportLab is
used to create the final two-page PDF.
"""

from __future__ import annotations

import html.parser
import io
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import winreg
except ImportError:  # The utility is Windows-only, but this keeps imports testable elsewhere.
    winreg = None  # type: ignore[assignment]

try:
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas as reportlab_canvas
except ImportError as reportlab_import_error:
    ImageReader = None
    stringWidth = None
    reportlab_canvas = None
    REPORTLAB_IMPORT_ERROR = reportlab_import_error
else:
    REPORTLAB_IMPORT_ERROR = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
except ImportError as pdf_reader_import_error:
    PdfReader = None
    PdfWriter = None
    PdfReadError = ValueError
    PDF_READER_IMPORT_ERROR = pdf_reader_import_error
else:
    PDF_READER_IMPORT_ERROR = None

try:
    import pypdfium2 as pdfium
except ImportError as pdfium_import_error:
    pdfium = None
    PDFIUM_IMPORT_ERROR = pdfium_import_error
else:
    PDFIUM_IMPORT_ERROR = None

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError as ocr_import_error:
    RapidOCR = None
    OCR_IMPORT_ERROR = ocr_import_error
else:
    OCR_IMPORT_ERROR = None


APP_TITLE = "Abaca Fiber Proforma & RR Packet"
APP_BACKGROUND = "#F8F9FA"
SIDEBAR_BACKGROUND = "#F1F3F5"
HEADING_FOREGROUND = "#1A202C"
LABEL_FOREGROUND = "#4A5568"
ENTRY_BORDER = "#E9ECEF"
ENTRY_WARNING_BACKGROUND = "#FFF3CD"
PRICE_FILE = Path(__file__).resolve().with_name("fiber_prices.txt")
OUTPUT_TEMPLATE_PATH = Path(__file__).resolve().with_name("OUTPUT FILE FORMAT.pdf")
DEFAULT_PRICES = {
    "S2": 95.00,
    "I": 89.00,
    "G": 72.00,
    "AD1": 92.00,
    "S3": 72.00,
}
REGISTRY_PATHS = (
    r"SOFTWARE\Tracker Software\PDFXEditor",
    r"SOFTWARE\WOW6432Node\Tracker Software\PDFXEditor",
)
DOWNLOAD_PAGE = "https://www.pdf-xchange.com/product/downloads"
ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
RR_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
RR_GRADES = ("S2", "I", "G", "AD1", "S3")
EDITOR_EXECUTABLE_NAMES = ("PXCEditor.exe", "PDFXEdit.exe")
_OCR_ENGINE: Any = None


def extract_dynamic_rr_weights(pdf_text_content: str) -> tuple[dict[str, float], list[str]]:
    """Extract net weights for each grade from text rows in a Magnera RR PDF.

    The value immediately preceding the bale-count marker ``40`` is treated as
    net weight. When the complete report is present, only its primary receiving
    report table is used; later receipts and evaluation tables repeat grade
    weights but do not represent additional delivered quantities.
    """
    report_start = re.search(
        r"\bABACA\s*FIBER\s*RECEIVING\s*REPORT\b",
        pdf_text_content,
        re.IGNORECASE,
    )
    if report_start is not None:
        report_body = pdf_text_content[report_start.end():]
        table_end = re.search(
            r"(?im)^\s*Total\s+[\d,]+(?:\.\d+)?\b",
            report_body,
        )
        if table_end is None:
            table_end = re.search(r"(?im)^\s*VAN\s+NO\s*:", report_body)
        if table_end is not None:
            pdf_text_content = report_body[:table_end.start()]
        else:
            pdf_text_content = report_body

    weights = {grade: 0.0 for grade in RR_GRADES}
    found_grades: set[str] = set()
    notes: list[str] = []
    grade_rows = re.compile(
        r"\b(?P<grade>AD1|S2|S3|I|G)\s+Fiber\b(?P<row>.*?)"
        r"(?=\b(?:AD1|S2|S3|I|G)\s+Fiber\b|\Z)",
        re.IGNORECASE | re.DOTALL,
    )
    weight_before_bale_count = re.compile(
        r"(?P<weight>\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)"
        r"\s+(?P<bales>\d{1,3})\s*$",
        re.MULTILINE,
    )

    for row_match in grade_rows.finditer(pdf_text_content):
        grade = row_match.group("grade").upper()
        found_grades.add(grade)
        weight_matches = list(weight_before_bale_count.finditer(row_match.group("row")))
        if not weight_matches:
            notes.append(
                f"{grade}: row found, but net weight before the bale count was not."
            )
            continue
        weight_match = weight_matches[-1]
        weight = float(weight_match.group("weight").replace(",", ""))
        if not math.isfinite(weight):
            notes.append(f"{grade}: extracted weight is too large; enter it manually.")
            continue
        weights[grade] += weight
        if not math.isfinite(weights[grade]):
            weights[grade] = 0.0
            notes.append(f"{grade}: combined weight is too large; enter it manually.")

    for grade in RR_GRADES:
        if grade not in found_grades:
            notes.append(f"{grade}: row not found; weight set to 0.")

    return weights, notes


def extract_rr_pdf_text(pdf_path: Path) -> str:
    """Extract searchable text from a receiving-report PDF."""
    if PdfReader is None:
        raise RuntimeError(
            "PDF text extraction requires pypdf. Install it with: "
            "py -m pip install pypdf"
        ) from PDF_READER_IMPORT_ERROR
    try:
        reader = PdfReader(str(pdf_path))
    except PdfReadError as error:
        raise ValueError(f"Could not read the receiving-report PDF: {error}") from error
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_dynamic_rr_metadata(pdf_text_content: str) -> dict[str, str]:
    """Extract RR number, purchase-order number, and billing date when present."""
    normalized_text = re.sub(r"\s+", " ", pdf_text_content)
    patterns = {
        "rr_number": (
            r"\bR\s*\.?\s*R\s*\.?\s*(?:N[O0]|NUMBER)\b"
            r"[\s:.,#-]*([A-Z0-9][A-Z0-9/-]*)"
        ),
        "po_number": (
            r"\bP\s*\.?\s*O\s*\.?\s*(?:N[O0]|NUMBER)\b"
            r"[\s:.,#-]*([A-Z0-9][A-Z0-9/-]*)"
        ),
        "billing_date": (
            r"\bDATE\b[\s:.,#-]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"
        ),
    }
    metadata: dict[str, str] = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, normalized_text, re.IGNORECASE)
        if match is None:
            continue
        value = match.group(1).strip(" .,:;")
        if field == "billing_date":
            parsed_date = None
            for date_format in ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y"):
                try:
                    parsed_date = datetime.strptime(value, date_format)
                    break
                except ValueError:
                    continue
            if parsed_date is not None:
                value = parsed_date.strftime("%m/%d/%Y")
        metadata[field] = value
    return metadata


def extract_sales_invoice_number(pdf_text_content: str) -> str:
    """Find a supplier's sales invoice number from OCR or searchable text."""
    normalized_text = re.sub(r"\s+", " ", pdf_text_content)
    patterns = (
        r"\b(?:SALES\s+)?INVOICE\s*(?:N[O0]|NUMBER|#)\.?\s*[:#.-]*"
        r"\s*([A-Z0-9][A-Z0-9/-]{1,})",
        r"\bS\.?\s*I\.?\s*(?:N[O0]|NUMBER|#)\.?\s*[:#.-]*"
        r"\s*([A-Z0-9][A-Z0-9/-]{1,})",
    )
    for pattern in patterns:
        match = re.search(pattern, normalized_text, re.IGNORECASE)
        if match:
            return match.group(1).strip(" .,:;")
    return ""


def _ocr_detections_to_text(detections: Any) -> str:
    """Reconstruct readable rows from RapidOCR's position-tagged detections."""
    positioned: list[tuple[float, float, float, str]] = []
    for detection in detections or []:
        if len(detection) < 2:
            continue
        points, text = detection[0], str(detection[1]).strip()
        if not text or not points:
            continue
        try:
            x_center = sum(float(point[0]) for point in points) / len(points)
            y_values = [float(point[1]) for point in points]
            y_center = sum(y_values) / len(y_values)
            height = max(y_values) - min(y_values)
        except (TypeError, ValueError, IndexError):
            continue
        positioned.append((y_center, x_center, height, text))

    positioned.sort(key=lambda item: (item[0], item[1]))
    lines: list[dict[str, Any]] = []
    for y_center, x_center, height, text in positioned:
        matching_line = next(
            (
                line
                for line in reversed(lines)
                if abs(y_center - line["center"])
                <= max(12.0, min(30.0, (height + line["height"]) * 0.42))
            ),
            None,
        )
        if matching_line is None:
            lines.append(
                {
                    "center": y_center,
                    "height": height,
                    "items": [(y_center, x_center, text)],
                }
            )
        else:
            matching_line["items"].append((y_center, x_center, text))
            matching_line["center"] = sum(
                item[0] for item in matching_line["items"]
            ) / len(matching_line["items"])
            matching_line["height"] = max(matching_line["height"], height)

    return "\n".join(
        " ".join(text for _, _, text in sorted(line["items"], key=lambda item: item[1]))
        for line in sorted(lines, key=lambda item: item["center"])
    )


def extract_rr_ocr_text(document_path: Path) -> str:
    """OCR an RR PDF or supported image locally using RapidOCR and PDFium."""
    global _OCR_ENGINE
    if RapidOCR is None:
        raise RuntimeError(
            "OCR package is missing. Install project dependencies with:\n"
            f'"{sys.executable}" -m pip install -r '
            f'"{Path(__file__).resolve().with_name("requirements.txt")}"'
        ) from OCR_IMPORT_ERROR
    if document_path.suffix.lower() == ".pdf" and pdfium is None:
        raise RuntimeError(
            "PDF image rendering package is missing. Install project dependencies "
            "with:\n"
            f'"{sys.executable}" -m pip install -r '
            f'"{Path(__file__).resolve().with_name("requirements.txt")}"'
        ) from PDFIUM_IMPORT_ERROR

    if _OCR_ENGINE is None:
        _OCR_ENGINE = RapidOCR()

    extracted_pages: list[str] = []
    if document_path.suffix.lower() == ".pdf":
        pdf_document = pdfium.PdfDocument(str(document_path))
        try:
            for page_index in range(len(pdf_document)):
                page = pdf_document[page_index]
                try:
                    bitmap = page.render(scale=2.5)
                    try:
                        detections, _ = _OCR_ENGINE(bitmap.to_numpy())
                    finally:
                        bitmap.close()
                finally:
                    page.close()
                page_text = _ocr_detections_to_text(detections)
                if page_text:
                    extracted_pages.append(page_text)
        finally:
            pdf_document.close()
    else:
        detections, _ = _OCR_ENGINE(str(document_path))
        page_text = _ocr_detections_to_text(detections)
        if page_text:
            extracted_pages.append(page_text)
    return "\n".join(extracted_pages)


def extract_rr_document_data(
    document_path: Path,
) -> tuple[dict[str, float], dict[str, str], list[str]]:
    """Extract weights and invoice metadata, using OCR when text is absent/incomplete."""
    text = ""
    notes: list[str] = []
    try:
        if document_path.suffix.lower() == ".pdf":
            text = extract_rr_pdf_text(document_path)
    except (OSError, ValueError, RuntimeError) as error:
        notes.append(f"Could not read the PDF text layer: {error}")

    weights, weight_notes = extract_dynamic_rr_weights(text)
    metadata = extract_dynamic_rr_metadata(text)
    grades_needing_ocr = {
        grade
        for grade in RR_GRADES
        if any(note.startswith(f"{grade}:") for note in weight_notes)
    }
    missing_metadata = any(
        field not in metadata for field in ("rr_number", "po_number")
    )

    if grades_needing_ocr or missing_metadata or not text.strip():
        try:
            ocr_text = extract_rr_ocr_text(document_path)
            if ocr_text:
                ocr_weights, _ = extract_dynamic_rr_weights(ocr_text)
                ocr_metadata = extract_dynamic_rr_metadata(ocr_text)
                for grade in RR_GRADES:
                    if grade in grades_needing_ocr:
                        weights[grade] = ocr_weights[grade]
                weight_notes = [
                    note
                    for note in weight_notes
                    if note.split(":", 1)[0] not in grades_needing_ocr
                ]
                verified_grades = {
                    grade
                    for grade in grades_needing_ocr
                    if ocr_weights[grade] > 0
                }
                for grade in sorted(grades_needing_ocr - verified_grades):
                    weight_notes.append(
                        f"{grade}: RR net weight could not be verified by OCR; "
                        "enter it manually."
                    )
                metadata.update(
                    {key: value for key, value in ocr_metadata.items() if not metadata.get(key)}
                )
                text = ocr_text
            else:
                notes.append("OCR found no readable text in the imported report.")
        except (OSError, ValueError, RuntimeError) as error:
            notes.append(f"OCR could not read the imported report: {error}")

    notes.extend(weight_notes)
    for field, label in (
        ("rr_number", "R.R. Number"),
        ("po_number", "P.O. Number"),
    ):
        if not metadata.get(field):
            notes.append(f"{label} was not detected; enter it manually.")
    if not metadata.get("billing_date"):
        notes.append("Billing date was not detected; enter it manually.")
    return weights, metadata, notes


def extract_sales_invoice_data(image_path: Path) -> tuple[str, dict[str, str], str]:
    """OCR supplier invoice number and any PO/date values present on its image."""
    try:
        text = extract_rr_ocr_text(image_path)
    except (OSError, ValueError, RuntimeError) as error:
        return "", {}, f"Sales invoice OCR failed: {error}"

    invoice_number = extract_sales_invoice_number(text)
    metadata = extract_dynamic_rr_metadata(text)
    note = "" if invoice_number else "Sales invoice number was not detected; enter it manually."
    return invoice_number, metadata, note


def create_root() -> tk.Tk:
    """Create a Tk root with optional native file-drop support."""
    if TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def _candidate_executables_from_registry() -> list[Path]:
    candidates: list[Path] = []
    if winreg is None:
        return candidates

    def add_install_location(value: Any) -> None:
        if not isinstance(value, str) or not value.strip():
            return
        location = Path(os.path.expandvars(value.strip().strip('"')))
        if location.name.lower() in {name.lower() for name in EDITOR_EXECUTABLE_NAMES}:
            candidates.append(location)
        else:
            candidates.extend(location / name for name in EDITOR_EXECUTABLE_NAMES)

    hives = (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE)
    views = (
        getattr(winreg, "KEY_WOW64_64KEY", 0),
        getattr(winreg, "KEY_WOW64_32KEY", 0),
    )
    for hive in hives:
        for registry_path in REGISTRY_PATHS:
            for view in views:
                try:
                    with winreg.OpenKey(
                        hive, registry_path, 0, winreg.KEY_READ | view
                    ) as key:
                        try:
                            default_value, _ = winreg.QueryValueEx(key, "")
                            add_install_location(default_value)
                        except OSError:
                            pass
                        value_index = 0
                        while True:
                            try:
                                name, value, value_type = winreg.EnumValue(
                                    key, value_index
                                )
                            except OSError:
                                break
                            value_index += 1
                            if value_type not in (
                                winreg.REG_SZ,
                                getattr(winreg, "REG_EXPAND_SZ", winreg.REG_SZ),
                            ) or not isinstance(value, str):
                                continue
                            if any(
                                marker in name.lower()
                                for marker in ("path", "dir", "folder", "location")
                            ):
                                add_install_location(value)
                except OSError:
                    continue

    # Some installs register the executable separately from the product key.
    for hive in hives:
        for executable_name in EDITOR_EXECUTABLE_NAMES:
            app_path_key = (
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"
                + executable_name
            )
            for view in views:
                try:
                    with winreg.OpenKey(
                        hive, app_path_key, 0, winreg.KEY_READ | view
                    ) as key:
                        value, _ = winreg.QueryValueEx(key, "")
                        if isinstance(value, str):
                            candidates.append(Path(os.path.expandvars(value.strip('"'))))
                except OSError:
                    continue

    return candidates


def find_pdfx_editor_in_folder(folder: Path) -> Path | None:
    """Find a supported editor executable in a folder or nearby subfolders."""
    folder = folder.expanduser().resolve()
    if not folder.is_dir():
        return None

    for executable_name in EDITOR_EXECUTABLE_NAMES:
        common_locations = (
            folder / executable_name,
            folder / "PDF Editor" / executable_name,
            folder / "PDF-XChange Editor" / executable_name,
            folder / "Tracker Software" / "PDF Editor" / executable_name,
            folder / "Tracker Software" / "PDF-XChange Editor" / executable_name,
        )
        for candidate in common_locations:
            if candidate.is_file():
                return candidate

    # Search a few levels only; the user-selected directory should be the
    # installation folder, not an entire drive or the Program Files root.
    for current, directories, filenames in os.walk(folder):
        relative_depth = len(Path(current).relative_to(folder).parts)
        if relative_depth >= 3:
            directories.clear()
        for executable_name in EDITOR_EXECUTABLE_NAMES:
            match = next(
                (
                    filename
                    for filename in filenames
                    if filename.lower() == executable_name.lower()
                ),
                None,
            )
            if match is not None:
                return Path(current) / match
    return None


def locate_pdfx_editor_manually(root: tk.Tk) -> Path | None:
    folder_name = filedialog.askdirectory(
        parent=root,
        title="Select the PDF-XChange Editor installation folder",
        mustexist=True,
    )
    if not folder_name:
        return None

    folder = Path(folder_name)
    editor_path = find_pdfx_editor_in_folder(folder)
    if editor_path is not None:
        return editor_path

    messagebox.showerror(
        "Editor executable not found",
        "PXCEditor.exe or PDFXEdit.exe was not found in the selected folder or its nearby "
        "subfolders.\n\nSelect the actual PDF-XChange Editor installation "
        "folder (the folder containing one of those executable files).",
        parent=root,
    )
    return None


def find_pdfx_editor() -> Path | None:
    """Return the installed PDF-XChange Editor executable, if it can be found."""
    registry_candidates = _candidate_executables_from_registry()
    candidates = list(registry_candidates)

    for variable in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(variable)
        if root:
            candidates.extend(
                (
                    Path(root) / "Tracker Software" / "PDF Editor" / "PXCEditor.exe",
                    Path(root) / "Tracker Software" / "PDF Editor" / "PDFXEdit.exe",
                    Path(root) / "Tracker Software" / "PDF-XChange Editor"
                    / "PXCEditor.exe",
                    Path(root) / "Tracker Software" / "PDF-XChange Editor"
                    / "PDFXEdit.exe",
                )
            )

    seen: set[str] = set()
    for candidate in candidates:
        normalized = os.path.normcase(os.path.abspath(str(candidate)))
        if normalized in seen:
            continue
        seen.add(normalized)
        if candidate.is_file():
            return candidate
    return None


class _DownloadLinkParser(html.parser.HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(
        self, tag: str, attrs: list[tuple[str, str | None]]
    ) -> None:
        if tag.lower() != "a":
            return
        href = dict(attrs).get("href")
        if href:
            self.links.append(href)


def find_official_installer_url() -> str:
    """Find an executable installer link published on the official product site."""
    product_page = "https://www.pdf-xchange.com/product/pdf-xchange-editor"
    for page_to_check in (DOWNLOAD_PAGE, product_page):
        request = urllib.request.Request(
            page_to_check, headers={"User-Agent": "FiberReportPacket/1.0"}
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            page_url = response.geturl()
            if not _is_pdfxchange_host(urllib.parse.urlparse(page_url).hostname):
                raise RuntimeError("The PDF-XChange download page redirected off-site.")
            page = response.read(8 * 1024 * 1024).decode("utf-8", errors="replace")

        parser = _DownloadLinkParser()
        parser.feed(page)
        for href in parser.links:
            url = urllib.parse.urljoin(page_url, href)
            parsed = urllib.parse.urlparse(url)
            if (
                parsed.scheme == "https"
                and _is_pdfxchange_host(parsed.hostname)
                and Path(parsed.path).suffix.lower() in {".exe", ".msi"}
            ):
                return url
    raise RuntimeError(
        "No direct installer link was published on the official download page."
    )


def _is_pdfxchange_host(hostname: str | None) -> bool:
    return bool(
        hostname
        and (
            hostname.lower() == "pdf-xchange.com"
            or hostname.lower().endswith(".pdf-xchange.com")
        )
    )


def download_official_installer(destination: Path, url: str) -> Path:
    """Download an installer only from an HTTPS PDF-XChange-owned host."""
    parsed_url = urllib.parse.urlparse(url)
    if parsed_url.scheme != "https" or not _is_pdfxchange_host(parsed_url.hostname):
        raise RuntimeError("Installer URLs must use HTTPS on a PDF-XChange-owned host.")
    request = urllib.request.Request(
        url, headers={"User-Agent": "FiberReportPacket/1.0"}
    )
    temporary_path = destination.with_name(destination.name + ".part")
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            final_url = urllib.parse.urlparse(response.geturl())
            if final_url.scheme != "https" or not _is_pdfxchange_host(
                final_url.hostname
            ):
                raise RuntimeError("The installer download redirected off-site.")
            content_length = response.headers.get("Content-Length")
            if content_length and int(content_length) > 1024 * 1024 * 1024:
                raise RuntimeError("The installer download exceeds the 1 GB limit.")

            total = 0
            with temporary_path.open("wb") as installer_file:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > 1024 * 1024 * 1024:
                        raise RuntimeError("The installer download exceeds the 1 GB limit.")
                    installer_file.write(chunk)

        with temporary_path.open("rb") as installer_file:
            signature = installer_file.read(8)
        if not (
            signature.startswith(b"MZ")
            or signature == bytes.fromhex("D0CF11E0A1B11AE1")
        ):
            raise RuntimeError("The downloaded file is not a recognizable Windows installer.")
        temporary_path.replace(destination)
        return destination
    except (OSError, ValueError, RuntimeError, urllib.error.URLError):
        temporary_path.unlink(missing_ok=True)
        raise


def install_prompt_and_exit(root: tk.Tk) -> bool:
    """Offer the official installer and launch it; return False to stop this run."""
    if not messagebox.askyesno(
        "PDF-XChange Editor required",
        "PDF-XChange Editor was not found. Download its official installer, "
        "save it locally, and launch it now?",
        parent=root,
    ):
        return False

    try:
        installer_url = find_official_installer_url()
    except (OSError, ValueError, RuntimeError, urllib.error.URLError) as error:
        messagebox.showerror(
            "Installer unavailable",
            f"Could not locate an installer link on the official site.\n\n{error}\n\n"
            f"You can download it manually from:\n{DOWNLOAD_PAGE}",
            parent=root,
        )
        webbrowser.open(DOWNLOAD_PAGE)
        return False

    suffix = Path(urllib.parse.urlparse(installer_url).path).suffix.lower()
    suggested_name = "PDF-XChange-Editor-Installer" + suffix
    downloads_directory = Path.home() / "Downloads"
    destination_name = filedialog.asksaveasfilename(
        parent=root,
        title="Save the official PDF-XChange Editor installer",
        initialdir=str(downloads_directory if downloads_directory.is_dir() else Path.home()),
        initialfile=suggested_name,
        defaultextension=suffix,
        filetypes=[("Windows installer", "*.exe *.msi"), ("All files", "*.*")],
    )
    if not destination_name:
        return False

    try:
        destination = Path(destination_name)
        if destination.suffix.lower() != suffix:
            destination = destination.with_suffix(suffix)
        installer_path = download_official_installer(destination, installer_url)
        os.startfile(str(installer_path))  # type: ignore[attr-defined]
    except (OSError, ValueError, RuntimeError, urllib.error.URLError) as error:
        messagebox.showerror(
            "Installer could not be launched",
            f"The installer could not be downloaded or started:\n\n{error}",
            parent=root,
        )
        return False

    messagebox.showinfo(
        "Installer launched",
        "The PDF-XChange Editor installer has been launched. Complete installation, "
        "then start this utility again.",
        parent=root,
    )
    return False


def load_prices() -> dict[str, float]:
    """Create the default price file if needed, then parse its grade/price pairs."""
    if not PRICE_FILE.exists():
        defaults = "".join(
            f"{grade}={price:.2f}\n" for grade, price in DEFAULT_PRICES.items()
        )
        PRICE_FILE.write_text(defaults, encoding="utf-8")

    prices = dict(DEFAULT_PRICES)
    for line_number, raw_line in enumerate(
        PRICE_FILE.read_text(encoding="utf-8-sig").splitlines(), start=1
    ):
        line = raw_line.split("#", 1)[0].strip()
        if not line:
            continue
        if "=" not in line:
            raise ValueError(f"{PRICE_FILE.name}, line {line_number}: expected GRADE=PRICE.")
        grade, raw_price = (part.strip() for part in line.split("=", 1))
        if not grade:
            raise ValueError(f"{PRICE_FILE.name}, line {line_number}: grade is empty.")
        try:
            price = float(raw_price)
        except ValueError as error:
            raise ValueError(
                f"{PRICE_FILE.name}, line {line_number}: {raw_price!r} is not a price."
            ) from error
        if not math.isfinite(price) or price < 0:
            raise ValueError(
                f"{PRICE_FILE.name}, line {line_number}: price must be finite and non-negative."
            )
        prices[grade] = price

    return prices


def parse_weight(raw_value: str) -> float:
    """Parse a non-negative weight; an empty field represents a zero-weight line."""
    text = raw_value.strip()
    if not text:
        return 0.0
    if "," in text and not re.fullmatch(
        r"\d{1,3}(?:,\d{3})+(?:\.\d*)?", text
    ):
        raise ValueError("use commas only as valid thousands separators")
    normalized = text.replace(",", "")
    if not re.fullmatch(r"(?:\d+(?:\.\d*)?|\.\d+)", normalized):
        raise ValueError("enter a non-negative decimal number")
    try:
        weight = float(normalized)
    except ValueError as error:
        raise ValueError("enter a number") from error
    if not math.isfinite(weight):
        raise ValueError("enter a finite number")
    if weight < 0:
        raise ValueError("weight cannot be negative")
    return weight


def _money(value: float) -> str:
    return f"PHP {value:,.2f}"


def build_report_pdf(
    output_path: Path,
    invoice_image: Path,
    output_template: Path,
    metadata: dict[str, str],
    weights: dict[str, float],
    prices: dict[str, float],
) -> None:
    if REPORTLAB_IMPORT_ERROR is not None:
        raise RuntimeError(
            "ReportLab is not installed. Install it with: py -m pip install reportlab"
        ) from REPORTLAB_IMPORT_ERROR
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError(
            "pypdf is not installed. Install it with: py -m pip install pypdf"
        ) from PDF_READER_IMPORT_ERROR
    if not output_template.is_file():
        raise FileNotFoundError(f"Output-form PDF not found: {output_template}")
    if not invoice_image.is_file():
        raise FileNotFoundError(f"Supplier invoice image not found: {invoice_image}")

    template_reader = PdfReader(str(output_template))
    if not template_reader.pages:
        raise ValueError("The output-form PDF does not contain any pages.")
    form_page = template_reader.pages[0]
    page_width = float(form_page.mediabox.width)
    page_height = float(form_page.mediabox.height)
    scale_x = page_width / 893
    scale_y = page_height / 1263

    def form_point(x: float, y_from_top: float) -> tuple[float, float]:
        return x * scale_x, page_height - y_from_top * scale_y

    def fitted_centered_text(
        drawing: Any, text: str, x: float, y: float, max_width: float, size: float = 8.5
    ) -> None:
        font_size = size
        while font_size > 5 and stringWidth(text, "Helvetica", font_size) > max_width:
            font_size -= 0.25
        drawing.setFont("Helvetica", font_size)
        drawing.drawCentredString(x, y, text)

    form_overlay = io.BytesIO()
    drawing = reportlab_canvas.Canvas(form_overlay, pagesize=(page_width, page_height))

    invoice_number_x, invoice_number_y = form_point(742, 356)
    fitted_centered_text(
        drawing,
        metadata.get("sales_invoice_number", ""),
        invoice_number_x,
        invoice_number_y,
        125 * scale_x,
    )
    for key, x, y, max_width in (
        ("billing_date", 744, 405, 125),
        ("po_number", 744, 428, 125),
        ("rr_number", 744, 451, 125),
    ):
        cell_x, cell_y = form_point(x, y)
        fitted_centered_text(
            drawing, metadata.get(key, ""), cell_x, cell_y, max_width * scale_x
        )

    column_centers = (130.5, 216.5, 312.5, 431, 545, 659, 764.5)
    row_baselines = (544, 567, 590, 613, 636)
    column_widths = (100, 58, 118, 100, 106, 100, 92)
    active_grades = [
        grade for grade in RR_GRADES if weights.get(grade, 0.0) > 0
    ]
    for grade, row_y in zip(active_grades, row_baselines):
        weight = weights[grade]
        price = prices[grade]
        values = (
            f"{weight:,.2f}",
            "KGS",
            f"{grade} Fiber",
            "PHP",
            f"{price:,.2f}",
            "PHP",
            f"{weight * price:,.2f}",
        )
        for value, center_x, column_width in zip(
            values, column_centers, column_widths
        ):
            x, y = form_point(center_x, row_y)
            fitted_centered_text(
                drawing, value, x, y, column_width * scale_x, size=7.5
            )

    total_x, total_y = form_point(column_centers[-1], 951)
    fitted_centered_text(
        drawing,
        f"{sum(weights[grade] * prices[grade] for grade in RR_GRADES):,.2f}",
        total_x,
        total_y,
        column_widths[-1] * scale_x,
        size=8,
    )
    drawing.save()
    form_overlay.seek(0)
    form_page.merge_page(PdfReader(form_overlay).pages[0])

    invoice_overlay = io.BytesIO()
    invoice_drawing = reportlab_canvas.Canvas(
        invoice_overlay, pagesize=(page_width, page_height)
    )
    margin = 24
    invoice_drawing.drawImage(
        ImageReader(str(invoice_image)),
        margin,
        margin,
        width=page_width - 2 * margin,
        height=page_height - 2 * margin,
        preserveAspectRatio=True,
        anchor="c",
        mask="auto",
    )
    invoice_drawing.save()
    invoice_overlay.seek(0)

    writer = PdfWriter()
    writer.add_page(form_page)
    writer.add_page(PdfReader(invoice_overlay).pages[0])
    writer.add_metadata(
        {
            "/Title": f"Proforma Invoice - RR {metadata['rr_number']}",
            "/Author": "TENEPRE ABACA TRADING",
        }
    )
    with output_path.open("wb") as output_file:
        writer.write(output_file)


class FiberPacketApp:
    def __init__(
        self,
        root: tk.Tk,
        editor_path: Path,
        rr_path: Path,
        invoice_path: Path,
        prices: dict[str, float],
        initial_weights: dict[str, float],
        rr_weight_notes: list[str],
        initial_metadata: dict[str, str],
    ) -> None:
        self.root = root
        self.editor_path = editor_path
        self.rr_path = rr_path
        self.invoice_path = invoice_path
        self.prices = prices
        self.initial_metadata = initial_metadata
        self.output_template_path = (
            OUTPUT_TEMPLATE_PATH if OUTPUT_TEMPLATE_PATH.is_file() else None
        )
        self.weight_vars = {
            grade: tk.StringVar(
                value=(
                    f"{initial_weights[grade]:g}"
                    if initial_weights.get(grade, 0.0) > 0
                    else ""
                )
            )
            for grade in RR_GRADES
        }
        self.rr_weight_notes = rr_weight_notes
        self.rr_ocr_input_copies: list[Path] = []
        self.amount_labels: dict[str, ttk.Label] = {}
        self.input_entries: dict[str, tk.Entry] = {}
        self.input_labels = {
            "metadata.po_number": "P.O. Number",
            "metadata.rr_number": "R.R. Number",
            "metadata.billing_date": "Billing Date",
            "metadata.sales_invoice_number": "Sales Invoice #",
            **{
                f"weight.{grade}": f"{grade} net weight"
                for grade in RR_GRADES
            },
        }
        self.missing_fields_var = tk.StringVar()
        self.total_var = tk.StringVar(value="PHP 0.00")
        self.invoice_path_var = tk.StringVar(value=str(invoice_path))
        self.output_template_var = tk.StringVar(
            value=str(self.output_template_path) if self.output_template_path else ""
        )
        self.preview_status_var = tk.StringVar(value="Select a document to preview.")

        self.root.title(APP_TITLE)
        self.root.geometry("1020x680")
        self.root.minsize(860, 580)
        self.root.configure(bg=APP_BACKGROUND)
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background=APP_BACKGROUND)
        style.configure("Sidebar.TFrame", background=SIDEBAR_BACKGROUND)
        style.configure(
            "TLabel",
            background=APP_BACKGROUND,
            foreground=LABEL_FOREGROUND,
        )
        style.configure(
            "App.TLabel",
            background=APP_BACKGROUND,
            foreground=LABEL_FOREGROUND,
        )
        style.configure(
            "Heading.App.TLabel",
            background=APP_BACKGROUND,
            foreground=HEADING_FOREGROUND,
        )
        style.configure(
            "Sidebar.TLabel",
            background=SIDEBAR_BACKGROUND,
            foreground=LABEL_FOREGROUND,
        )
        style.configure(
            "Heading.Sidebar.TLabel",
            background=SIDEBAR_BACKGROUND,
            foreground=HEADING_FOREGROUND,
        )
        style.configure(
            "App.TLabelframe",
            background=APP_BACKGROUND,
            bordercolor=ENTRY_BORDER,
        )
        style.configure(
            "App.TLabelframe.Label",
            background=APP_BACKGROUND,
            foreground=HEADING_FOREGROUND,
        )
        style.configure(
            "App.TButton",
            background="#E2E8F0",
            foreground=HEADING_FOREGROUND,
            relief="flat",
            borderwidth=0,
            padding=(10, 7),
        )
        style.map(
            "App.TButton",
            background=[("active", "#CBD5E1"), ("pressed", "#CBD5E1")],
            foreground=[("disabled", LABEL_FOREGROUND)],
        )
        style.configure(
            "Primary.TButton",
            background="#1E3A8A",
            foreground="#FFFFFF",
            font=("Helvetica", 11, "bold"),
            relief="flat",
            borderwidth=0,
            padding=(10, 9),
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1E3A8A"), ("pressed", "#172554")],
            foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")],
        )
        style.configure(
            "Sidebar.TEntry",
            fieldbackground=SIDEBAR_BACKGROUND,
            foreground=HEADING_FOREGROUND,
            bordercolor=ENTRY_BORDER,
        )
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_invoice_path(self.invoice_path)
        for variable in self.weight_vars.values():
            variable.trace_add("write", self._refresh_summary)
            variable.trace_add("write", self._refresh_input_validation)
        for variable in self.metadata_vars.values():
            variable.trace_add("write", self._refresh_input_validation)
        self._refresh_summary()
        self._refresh_input_validation()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=18, style="App.TFrame")
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        self.left_panel = ttk.Frame(
            container, padding=(0, 0, 16, 0), style="App.TFrame"
        )
        self.left_panel.grid(row=0, column=0, sticky="nsew")
        right = ttk.Frame(container, padding=16, style="Sidebar.TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        ttk.Label(
            self.left_panel,
            text="Abaca Fiber Billing",
            font=("Segoe UI", 18, "bold"),
            style="Heading.App.TLabel",
        ).pack(anchor="w", pady=(0, 14))
        ttk.Label(
            self.left_panel,
            textvariable=self.missing_fields_var,
            foreground="#856404",
            background=APP_BACKGROUND,
            wraplength=570,
            justify="left",
        ).pack(anchor="w", pady=(0, 10))
        metadata_frame = ttk.LabelFrame(
            self.left_panel,
            text="Invoice metadata",
            padding=12,
            style="App.TLabelframe",
        )
        metadata_frame.pack(fill="x", pady=(0, 14))
        metadata_frame.columnconfigure(1, weight=1)
        self.metadata_vars = {
            "po_number": tk.StringVar(
                value=self.initial_metadata.get("po_number", "")
            ),
            "rr_number": tk.StringVar(
                value=self.initial_metadata.get("rr_number", "")
            ),
            "billing_date": tk.StringVar(
                value=self.initial_metadata.get("billing_date", "")
            ),
            "sales_invoice_number": tk.StringVar(
                value=self.initial_metadata.get("sales_invoice_number", "")
            ),
        }
        fields = (
            ("P.O. Number", "po_number"),
            ("R.R. Number", "rr_number"),
            ("Billing Date (MM/DD/YYYY)", "billing_date"),
            ("Sales Invoice #", "sales_invoice_number"),
        )
        for row, (label, key) in enumerate(fields):
            ttk.Label(metadata_frame, text=label).grid(
                row=row, column=0, sticky="w", padx=(0, 12), pady=5
            )
            entry = tk.Entry(
                metadata_frame,
                textvariable=self.metadata_vars[key],
                relief="flat",
                bd=1,
                highlightthickness=1,
                highlightbackground=ENTRY_BORDER,
                highlightcolor=ENTRY_BORDER,
                background="#FFFFFF",
            )
            entry.grid(
                row=row, column=1, sticky="ew", pady=5
            )
            self.input_entries[f"metadata.{key}"] = entry

        weights_frame = ttk.LabelFrame(
            self.left_panel,
            text="Fiber quantities & live billing",
            padding=12,
            style="App.TLabelframe",
        )
        weights_frame.pack(fill="x", pady=(0, 14))
        headings = ("Grade", "Net Weight (KGS)", "Price / Kilo", "Line Amount")
        for column, heading in enumerate(headings):
            ttk.Label(weights_frame, text=heading, font=("Segoe UI", 9, "bold")).grid(
                row=0, column=column, sticky="w", padx=5, pady=(0, 8)
            )
        for row, grade in enumerate(RR_GRADES, start=1):
            ttk.Label(weights_frame, text=grade).grid(
                row=row, column=0, sticky="w", padx=5, pady=5
            )
            weight_entry = tk.Entry(
                weights_frame,
                textvariable=self.weight_vars[grade],
                width=16,
                relief="flat",
                bd=1,
                highlightthickness=1,
                highlightbackground=ENTRY_BORDER,
                highlightcolor=ENTRY_BORDER,
                background="#FFFFFF",
            )
            weight_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=4)
            self.input_entries[f"weight.{grade}"] = weight_entry
            ttk.Label(
                weights_frame,
                text=f"PHP {self.prices[grade]:,.2f}",
                style="App.TLabel",
            ).grid(row=row, column=2, sticky="e", padx=5, pady=4)
            amount_label = ttk.Label(
                weights_frame,
                text="PHP 0.00",
                anchor="e",
                style="App.TLabel",
            )
            amount_label.grid(row=row, column=3, sticky="e", padx=5, pady=4)
            self.amount_labels[grade] = amount_label
        weights_frame.columnconfigure(1, weight=1)
        weights_frame.columnconfigure(3, weight=1)
        self.rr_notes_label = ttk.Label(
            weights_frame,
            wraplength=570,
            style="App.TLabel",
        )
        self.rr_notes_label.grid(
            row=len(RR_GRADES) + 1,
            column=0,
            columnspan=4,
            sticky="w",
            padx=5,
            pady=(8, 0),
        )
        self._refresh_rr_notes()

        total_box = ttk.LabelFrame(
            self.left_panel,
            text="Grand total billing amount",
            padding=12,
            style="App.TLabelframe",
        )
        total_box.pack(fill="x", pady=(0, 14))
        ttk.Label(
            total_box,
            textvariable=self.total_var,
            font=("Segoe UI", 16, "bold"),
            style="Heading.App.TLabel",
        ).pack(anchor="e")

        ttk.Button(
            self.left_panel,
            text="FINALIZE & EXPORT REPORT PACKET",
            command=self.finalize_report,
            style="Primary.TButton",
        ).pack(fill="x", ipady=8, pady=(6, 0))
        tk.Label(
            self.left_panel,
            text="Engine Core Designed by: RAVEN TENEPRE",
            font=("Helvetica", 8, "italic"),
            fg="#A0AEC0",
            bg=APP_BACKGROUND,
        ).pack(side="bottom", anchor="w", pady=(20, 0))

        ttk.Label(
            right,
            text="Document preview",
            font=("Segoe UI", 14, "bold"),
            style="Heading.Sidebar.TLabel",
        ).pack(
            anchor="w", pady=(0, 8)
        )
        ttk.Label(
            right,
            text="PDF-XChange Editor opens each document in its own window.",
            wraplength=300,
            style="Sidebar.TLabel",
        ).pack(anchor="w", pady=(0, 14))
        ttk.Button(
            right,
            text="Preview Receiving Report",
            command=lambda: self.open_preview(self.rr_path),
            style="App.TButton",
        ).pack(fill="x", pady=4)
        ttk.Button(
            right,
            text="Run PDF-XChange OCR (page 1)",
            command=self.run_pdfxchange_ocr,
            style="App.TButton",
        ).pack(fill="x", pady=4)
        ttk.Button(
            right,
            text="Import saved OCR PDF",
            command=self.import_ocr_result,
            style="App.TButton",
        ).pack(fill="x", pady=4)
        ttk.Label(
            right,
            text="Proforma output format PDF",
            style="Sidebar.TLabel",
        ).pack(
            anchor="w", pady=(10, 4)
        )
        self.output_template_entry = ttk.Entry(
            right,
            textvariable=self.output_template_var,
            state="readonly",
            style="Sidebar.TEntry",
        )
        self.output_template_entry.pack(fill="x", pady=(0, 4))
        ttk.Button(
            right,
            text="Choose output format PDF",
            command=self.choose_output_template,
            style="App.TButton",
        ).pack(fill="x", pady=4)
        ttk.Label(
            right,
            text="Supplier sales invoice file",
            style="Sidebar.TLabel",
        ).pack(
            anchor="w", pady=(14, 4)
        )
        self.invoice_entry = ttk.Entry(
            right,
            textvariable=self.invoice_path_var,
            state="readonly",
            style="Sidebar.TEntry",
        )
        self.invoice_entry.pack(fill="x", pady=(0, 6))
        ttk.Button(
            right,
            text="Choose invoice image",
            command=self.choose_invoice,
            style="App.TButton",
        ).pack(
            fill="x", pady=4
        )
        ttk.Button(
            right,
            text="Preview Sales Invoice",
            command=lambda: self.open_preview(self.invoice_path),
            style="App.TButton",
        ).pack(fill="x", pady=4)

        drop_text = (
            "Drag an invoice image here to replace the selected file."
            if DND_FILES is not None
            else "Optional drag-and-drop requires tkinterdnd2; use the chooser above."
        )
        self.drop_label = ttk.Label(
            right,
            text=drop_text,
            relief="ridge",
            padding=14,
            wraplength=280,
            anchor="center",
            style="Sidebar.TLabel",
        )
        self.drop_label.pack(fill="x", pady=(14, 8))
        if DND_FILES is not None:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_invoice_drop)

        ttk.Separator(right).pack(fill="x", pady=10)
        ttk.Label(
            right,
            textvariable=self.preview_status_var,
            wraplength=300,
            style="Sidebar.TLabel",
        ).pack(anchor="w")
        self.root.deiconify()

    def _refresh_summary(self, *_: Any) -> None:
        line_amounts: dict[str, float] = {}
        invalid: list[str] = []
        for grade, variable in self.weight_vars.items():
            try:
                weight = parse_weight(variable.get())
            except ValueError:
                self.amount_labels[grade].configure(text="Invalid weight")
                invalid.append(grade)
                continue
            amount = weight * self.prices[grade]
            if not math.isfinite(amount):
                self.amount_labels[grade].configure(text="Amount too large")
                invalid.append(grade)
                continue
            line_amounts[grade] = amount
            self.amount_labels[grade].configure(text=_money(amount))

        if invalid:
            self.total_var.set(f"Correct weight for: {', '.join(invalid)}")
        else:
            self.total_var.set(_money(sum(line_amounts.values())))

    def _refresh_input_validation(self, *_: Any) -> None:
        needs_review: list[str] = []
        invalid_fields: set[str] = set()

        for key, variable in self.metadata_vars.items():
            field_key = f"metadata.{key}"
            value = variable.get().strip()
            invalid = not value
            if key == "billing_date" and value:
                try:
                    parsed_date = datetime.strptime(value, "%m/%d/%Y")
                    invalid = parsed_date.strftime("%m/%d/%Y") != value
                except ValueError:
                    invalid = True
            if invalid:
                invalid_fields.add(field_key)
                if key == "billing_date" and value:
                    needs_review.append("Billing Date must use MM/DD/YYYY")
                else:
                    needs_review.append(self.input_labels[field_key])

        for grade, variable in self.weight_vars.items():
            field_key = f"weight.{grade}"
            value = variable.get().strip()
            invalid = not value
            if value:
                try:
                    parse_weight(value)
                except ValueError:
                    invalid = True
                    needs_review.append(f"{grade} net weight is invalid")
            elif not value:
                needs_review.append(f"{grade} net weight is blank")
            if invalid:
                invalid_fields.add(field_key)

        for field_key, entry in self.input_entries.items():
            is_invalid = field_key in invalid_fields
            entry.configure(
                background=(
                    ENTRY_WARNING_BACKGROUND if is_invalid else "#FFFFFF"
                ),
                highlightbackground=ENTRY_BORDER,
                highlightcolor=ENTRY_BORDER,
            )

        if needs_review:
            self.missing_fields_var.set(
                "Review highlighted fields: "
                + ", ".join(needs_review)
                + "\nBlank weights are omitted from the PDF; enter 0 to confirm "
                "a grade was not delivered."
            )
        else:
            self.missing_fields_var.set("All input fields have values.")

    def _refresh_rr_notes(self) -> None:
        if self.rr_weight_notes:
            self.rr_notes_label.configure(
                text="RR import note: " + " ".join(self.rr_weight_notes),
                foreground="#8A5A00",
            )
        else:
            self.rr_notes_label.configure(
                text="Net weights imported from the RR; verify and edit as needed.",
                foreground="#405467",
            )

    def run_pdfxchange_ocr(self) -> None:
        if self.rr_path.suffix.lower() != ".pdf":
            messagebox.showerror(
                "PDF required",
                "PDF-XChange OCR can only be started for a receiving-report PDF.",
                parent=self.root,
            )
            return

        temporary_copy: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix="fiber_rr_ocr_input_",
                suffix=".pdf",
                delete=False,
            ) as temporary_file:
                temporary_copy = Path(temporary_file.name)
            shutil.copy2(self.rr_path, temporary_copy)
            subprocess.Popen(
                [
                    str(self.editor_path),
                    "/ocr",
                    "/pages:1",
                    "/accuracy",
                    "high",
                    "/lang",
                    "en",
                    "/type",
                    "search",
                    str(temporary_copy),
                ],
                close_fds=True,
            )
        except OSError as error:
            if temporary_copy is not None:
                temporary_copy.unlink(missing_ok=True)
            messagebox.showerror(
                "PDF-XChange OCR could not be started",
                f"The page-1 OCR request could not be opened:\n\n{error}",
                parent=self.root,
            )
            return

        assert temporary_copy is not None
        self.rr_ocr_input_copies.append(temporary_copy)
        self.preview_status_var.set(
            "OCR opened on a temporary copy. Save the searchable OCR result, "
            "then import it here."
        )
        messagebox.showinfo(
            "Complete PDF-XChange OCR",
            "PDF-XChange was opened with the page-1 OCR request. In the OCR dialog, "
            "use Current: 1, English, High accuracy, and Searchable Image. Leave "
            "'Ignore existing text on page' unchecked. Complete OCR, then use Save "
            "As to save the resulting searchable PDF. Return here and click "
            "'Import saved OCR PDF'. The original receiving report is not modified.",
            parent=self.root,
        )

    def import_ocr_result(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Select the PDF-XChange OCR result",
            filetypes=[("PDF document", "*.pdf"), ("All files", "*.*")],
        )
        if not selected:
            return
        result_path = Path(selected).expanduser().resolve()
        if not result_path.is_file() or result_path.suffix.lower() != ".pdf":
            messagebox.showerror(
                "Invalid OCR result",
                "Select an existing searchable PDF created by PDF-XChange OCR.",
                parent=self.root,
            )
            return
        if result_path == self.rr_path.resolve():
            messagebox.showerror(
                "Original report selected",
                "Select the PDF-XChange OCR result, not the original receiving report.",
                parent=self.root,
            )
            return

        try:
            weights, metadata, notes = extract_rr_document_data(result_path)
        except (OSError, ValueError, RuntimeError) as error:
            messagebox.showerror(
                "OCR result could not be read",
                f"The selected OCR PDF could not be processed:\n\n{error}",
                parent=self.root,
            )
            return

        if not any(weight > 0 for weight in weights.values()):
            detail = " ".join(notes) or "No fiber net weights were found."
            messagebox.showwarning(
                "Net weights not detected",
                f"No fiber weights were imported; the current entries were left unchanged.\n\n{detail}",
                parent=self.root,
            )
            return

        for grade in RR_GRADES:
            self.weight_vars[grade].set(
                f"{weights[grade]:g}" if weights[grade] > 0 else ""
            )
        for field in ("po_number", "rr_number", "billing_date"):
            if metadata.get(field):
                self.metadata_vars[field].set(metadata[field])

        self.rr_weight_notes = notes
        self._refresh_rr_notes()
        self.preview_status_var.set(
            f"Imported OCR data from {result_path.name}."
        )
        if notes:
            messagebox.showwarning(
                "OCR imported with notes",
                "The detected weights and metadata were imported. Please review:\n\n"
                + "\n".join(notes),
                parent=self.root,
            )
        else:
            messagebox.showinfo(
                "OCR data imported",
                "The detected net weights and metadata were copied into the input fields.",
                parent=self.root,
            )

    def _on_close(self) -> None:
        cleanup_errors: list[str] = []
        for temporary_copy in self.rr_ocr_input_copies:
            try:
                temporary_copy.unlink(missing_ok=True)
            except OSError as error:
                cleanup_errors.append(f"{temporary_copy}: {error}")
        if cleanup_errors:
            messagebox.showwarning(
                "Temporary OCR copy could not be removed",
                "Close the PDF-XChange OCR document and delete this temporary file:\n\n"
                + "\n".join(cleanup_errors),
                parent=self.root,
            )
        self.root.destroy()

    def open_preview(self, document_path: Path) -> None:
        if not document_path.is_file():
            messagebox.showerror(
                "Document unavailable",
                f"The selected document no longer exists:\n{document_path}",
                parent=self.root,
            )
            return
        try:
            subprocess.Popen(
                [str(self.editor_path), str(document_path)],
                close_fds=True,
            )
        except OSError as error:
            messagebox.showerror(
                "Preview could not be opened",
                f"PDF-XChange Editor could not open the document:\n\n{error}",
                parent=self.root,
            )
            self.preview_status_var.set(f"Preview failed: {document_path.name}")
            return
        self.preview_status_var.set(f"Opened in PDF-XChange: {document_path.name}")

    def choose_invoice(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Select supplier Sales Invoice image",
            filetypes=[
                ("Invoice images", "*.jpg *.jpeg *.png"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            try:
                self._set_invoice_path(Path(selected))
            except ValueError as error:
                messagebox.showerror("Invalid invoice image", str(error), parent=self.root)

    def choose_output_template(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Select the proforma output-format PDF",
            initialdir=str(OUTPUT_TEMPLATE_PATH.parent),
            filetypes=[("PDF document", "*.pdf"), ("All files", "*.*")],
        )
        if not selected:
            return
        path = Path(selected).expanduser().resolve()
        if not path.is_file() or path.suffix.lower() != ".pdf":
            messagebox.showerror(
                "Invalid output format",
                "Select an existing output-format PDF.",
                parent=self.root,
            )
            return
        self.output_template_path = path
        self.output_template_var.set(str(path))

    def _on_invoice_drop(self, event: Any) -> None:
        try:
            dropped_paths = self.root.tk.splitlist(event.data)
            if len(dropped_paths) != 1:
                raise ValueError("Drop exactly one invoice image.")
            self._set_invoice_path(Path(dropped_paths[0]))
        except (OSError, ValueError) as error:
            messagebox.showerror("Invalid dropped file", str(error), parent=self.root)

    def _set_invoice_path(self, path: Path) -> None:
        path = path.expanduser().resolve()
        if not path.is_file():
            raise ValueError("The selected invoice image does not exist.")
        if path.suffix.lower() not in ALLOWED_IMAGE_EXTENSIONS:
            raise ValueError("Choose a JPEG or PNG invoice image for the PDF packet.")
        self.invoice_path = path
        self.invoice_path_var.set(str(path))
        self.preview_status_var.set("Reading sales invoice fields locally...")
        self.root.update_idletasks()
        invoice_number, invoice_metadata, invoice_note = extract_sales_invoice_data(path)
        self.metadata_vars["sales_invoice_number"].set(invoice_number)
        for field in ("po_number", "billing_date"):
            if not self.metadata_vars[field].get().strip() and invoice_metadata.get(field):
                self.metadata_vars[field].set(invoice_metadata[field])
        self.preview_status_var.set(
            invoice_note or f"Invoice selected; detected invoice #: {invoice_number}"
        )

    def finalize_report(self) -> None:
        metadata = {key: variable.get().strip() for key, variable in self.metadata_vars.items()}
        missing = [
            label
            for key, label in (
                ("po_number", "P.O. Number"),
                ("rr_number", "R.R. Number"),
                ("billing_date", "Billing Date"),
            )
            if not metadata[key]
        ]
        if missing:
            messagebox.showerror(
                "Required information",
                "Enter the following field(s): " + ", ".join(missing),
                parent=self.root,
            )
            return
        try:
            parsed_date = datetime.strptime(metadata["billing_date"], "%m/%d/%Y")
            if parsed_date.strftime("%m/%d/%Y") != metadata["billing_date"]:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Invalid billing date",
                "Enter the billing date in MM/DD/YYYY format.",
                parent=self.root,
            )
            return

        weights: dict[str, float] = {}
        for grade, variable in self.weight_vars.items():
            try:
                weights[grade] = parse_weight(variable.get())
            except ValueError as error:
                messagebox.showerror(
                    "Invalid net weight",
                    f"{grade} weight must be a non-negative number ({error}).",
                    parent=self.root,
                )
                return
            if not math.isfinite(weights[grade] * self.prices[grade]):
                messagebox.showerror(
                    "Invalid net weight",
                    f"{grade} weight is too large to calculate a finite line amount.",
                    parent=self.root,
                )
                return

        if not math.isfinite(
            sum(weights[grade] * self.prices[grade] for grade in weights)
        ):
            messagebox.showerror(
                "Invalid net weights",
                "The combined billing amount is too large to represent.",
                parent=self.root,
            )
            return

        if not self.invoice_path.is_file():
            messagebox.showerror(
                "Invoice image required",
                "Select a valid supplier Sales Invoice image before exporting.",
                parent=self.root,
            )
            return
        if self.output_template_path is None or not self.output_template_path.is_file():
            messagebox.showerror(
                "Output format required",
                "Choose the proforma output-format PDF before exporting.",
                parent=self.root,
            )
            return
        safe_rr_number = re.sub(
            r'[<>:"/\\|?*\x00-\x1f]', "_", metadata["rr_number"]
        ).strip(" .")
        if not safe_rr_number:
            messagebox.showerror(
                "Invalid R.R. Number",
                "The R.R. Number must contain at least one filename-safe character.",
                parent=self.root,
            )
            return
        suggested_name = f"Finalized_Proforma_RR_{safe_rr_number}.pdf"
        output_name = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save finalized report packet",
            initialdir=str(self.invoice_path.parent),
            initialfile=suggested_name,
            defaultextension=".pdf",
            filetypes=[("PDF document", "*.pdf")],
        )
        if not output_name:
            return

        output_path = Path(output_name)
        if output_path.suffix.lower() != ".pdf":
            output_path = output_path.with_suffix(".pdf")
        if output_path.resolve() in {
            self.output_template_path.resolve(),
            self.invoice_path.resolve(),
        }:
            messagebox.showerror(
                "Invalid output path",
                "The output packet cannot overwrite the output form or invoice image.",
                parent=self.root,
            )
            return
        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix="fiber_packet_",
                suffix=".pdf",
                dir=str(output_path.parent),
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
            build_report_pdf(
                temporary_path,
                self.invoice_path,
                self.output_template_path,
                metadata,
                weights,
                self.prices,
            )
            temporary_path.replace(output_path)
        except Exception as error:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)
            messagebox.showerror(
                "PDF export failed",
                f"The report packet could not be created:\n\n{error}",
                parent=self.root,
            )
            return

        messagebox.showinfo(
            "Report packet created",
            f"Saved the finalized two-page report packet to:\n{output_path}",
            parent=self.root,
        )


def _select_import_file(
    root: tk.Tk,
    title: str,
    filetypes: list[tuple[str, str]],
    *,
    require_pdf: bool = False,
    allowed_extensions: tuple[str, ...] = (),
) -> Path | None:
    while True:
        selected = filedialog.askopenfilename(
            parent=root,
            title=title,
            filetypes=filetypes,
        )
        if not selected:
            if messagebox.askretrycancel(
                "No File Selected",
                "You did not select a file.\n\nWould you like to try again?",
                parent=root,
            ):
                continue
            return None

        file_path = selected
        if require_pdf and not file_path.lower().endswith(".pdf"):
            messagebox.showerror(
                "Invalid File Format",
                "The imported file is incorrect.\n"
                "Please select a valid digital or scanned PDF Receiving Report.",
                parent=root,
            )
            file_path = ""
            continue

        path = Path(file_path)
        if not path.is_file():
            messagebox.showerror(
                "File Unavailable",
                f"The selected file does not exist or cannot be accessed:\n{path}",
                parent=root,
            )
            file_path = ""
            continue

        if allowed_extensions and path.suffix.lower() not in allowed_extensions:
            messagebox.showerror(
                "Invalid File Format",
                "The selected file format is not supported. Please choose a "
                "valid supplier Sales Invoice image.",
                parent=root,
            )
            file_path = ""
            continue

        return path


def main() -> int:
    root = create_root()
    root.withdraw()
    root.title(APP_TITLE)

    if os.name != "nt":
        messagebox.showerror(
            "Windows required",
            "This utility requires Windows and PDF-XChange Editor.",
            parent=root,
        )
        root.destroy()
        return 1

    editor_path = find_pdfx_editor()
    if editor_path is None:
        locate_manually = messagebox.askyesno(
            "PDF-XChange Editor not detected",
            "The program could not find PDF-XChange Editor automatically.\n\n"
            "Would you like to browse for its installation folder?",
            parent=root,
        )
        if locate_manually:
            editor_path = locate_pdfx_editor_manually(root)
        if editor_path is None:
            install_prompt_and_exit(root)
            root.destroy()
            return 0

    if REPORTLAB_IMPORT_ERROR is not None:
        requirements_path = Path(__file__).resolve().with_name("requirements.txt")
        install_command = (
            f'"{sys.executable}" -m pip install -r "{requirements_path}"'
        )
        messagebox.showerror(
            "Required packages are missing",
            "The PDF packet dependencies are not installed in the Python "
            "environment running this program.\n\nInstall them with:\n"
            f"{install_command}",
            parent=root,
        )
        root.destroy()
        return 1

    try:
        prices = load_prices()
    except (OSError, ValueError) as error:
        messagebox.showerror("Fiber prices unavailable", str(error), parent=root)
        root.destroy()
        return 1

    rr_path = _select_import_file(
        root,
        "Step 1 of 2 — Select Abaca Fiber Receiving Report (RR)",
        [
            ("PDF receiving report", "*.pdf"),
            ("All files", "*.*"),
        ],
        require_pdf=True,
    )
    if rr_path is None:
        root.destroy()
        return 0

    root.title(f"{APP_TITLE} — Importing documents")
    root.geometry("540x145")
    startup_status = ttk.Label(
        root,
        text=(
            "Receiving report selected.\n"
            "Opening the report, then choose the supplier sales invoice."
        ),
        padding=24,
        justify="center",
        wraplength=480,
    )
    startup_status.pack(fill="both", expand=True)
    root.deiconify()
    root.update_idletasks()

    # Preview the RR immediately, before asking for the supplier invoice.
    try:
        subprocess.Popen([str(editor_path), str(rr_path)], close_fds=True)
    except OSError as error:
        messagebox.showerror(
            "Receiving report preview failed",
            f"PDF-XChange Editor could not open the receiving report:\n\n{error}",
            parent=root,
        )

    invoice_path = _select_import_file(
        root,
        "Step 2 of 2 — Select supplier Sales Invoice image",
        [
            ("Invoice images", "*.jpg *.jpeg *.png"),
            ("All files", "*.*"),
        ],
        allowed_extensions=tuple(ALLOWED_IMAGE_EXTENSIONS),
    )
    if invoice_path is None:
        root.destroy()
        return 0
    if not invoice_path.is_file() or invoice_path.suffix.lower() not in ALLOWED_IMAGE_EXTENSIONS:
        messagebox.showerror(
            "Invalid sales invoice",
            "Select an existing JPEG or PNG Sales Invoice image.",
            parent=root,
        )
        root.destroy()
        return 1

    startup_status.configure(
        text="Reading the receiving report and supplier invoice with local OCR..."
    )
    root.update_idletasks()
    try:
        initial_weights, initial_metadata, rr_weight_notes = (
            extract_rr_document_data(rr_path)
        )
    except (OSError, ValueError, RuntimeError) as error:
        initial_weights = {grade: 0.0 for grade in RR_GRADES}
        initial_metadata = {}
        rr_weight_notes = [f"Could not scrape the receiving report: {error}"]

    startup_status.destroy()
    FiberPacketApp(
        root,
        editor_path,
        rr_path,
        invoice_path,
        prices,
        initial_weights,
        rr_weight_notes,
        initial_metadata,
    )
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
