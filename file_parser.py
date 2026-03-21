import os
import re
from io import BytesIO

from openpyxl import load_workbook

TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(os.getenv("USERNAME", "")),
]

DAY_MAP = {
    "mon": "Monday", "monday": "Monday",
    "tue": "Tuesday", "tues": "Tuesday", "tuesday": "Tuesday",
    "wed": "Wednesday", "wednesday": "Wednesday",
    "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday", "thursday": "Thursday",
    "fri": "Friday", "friday": "Friday",
    "sat": "Saturday", "saturday": "Saturday",
    "sun": "Sunday", "sunday": "Sunday",
}

VALID_TYPES = {"content", "email", "message"}


def parse_file(file_storage):
    """
    Parse an uploaded file and return a list of employee dicts.
    Supports .xlsx, .pdf, .png, .jpg, .jpeg.

    Each returned dict: {name, content_types: [], working_days: [], shift: int|None}
    Raises ValueError on parse failure.
    """
    filename = file_storage.filename.lower()
    file_bytes = file_storage.read()

    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        return parse_excel(file_bytes)
    elif filename.endswith(".pdf"):
        return parse_pdf(file_bytes)
    elif filename.endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
        return parse_image(file_bytes)
    else:
        raise ValueError(f"Unsupported file type: {filename}")


def parse_excel(file_bytes):
    wb = load_workbook(filename=BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active

    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([str(c).strip() if c is not None else "" for c in row])
    wb.close()

    if len(rows) < 2:
        raise ValueError("Excel file has no data rows.")

    return _parse_table(rows)


def parse_pdf(file_bytes):
    try:
        import pdfplumber
    except ImportError:
        raise ValueError("pdfplumber is not installed. Run: pip install pdfplumber")

    rows = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        cleaned = [str(c).strip() if c else "" for c in row]
                        rows.append(cleaned)
            else:
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        parts = re.split(r"\s{2,}|\t", line.strip())
                        if len(parts) >= 3:
                            rows.append(parts)

    if len(rows) < 2:
        raise ValueError("Could not extract enough data from the PDF.")

    return _parse_table(rows)


def parse_image(file_bytes):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise ValueError("pytesseract and Pillow are required. Run: pip install pytesseract Pillow")

    for path in TESSERACT_PATHS:
        if os.path.isfile(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break

    img = Image.open(BytesIO(file_bytes))
    text = pytesseract.image_to_string(img)

    if not text or not text.strip():
        raise ValueError("Could not extract any text from the image. Ensure Tesseract OCR is installed.")

    rows = []
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = re.split(r"\s{2,}|\t|\|", line)
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) >= 3:
            rows.append(parts)

    if len(rows) < 2:
        raise ValueError(
            "Could not parse enough rows from the image. "
            "Ensure the image contains a clear table with columns: "
            "Name, Shift, Working Days, Product Type."
        )

    return _parse_table(rows)


def _parse_table(rows):
    """
    Given a list of row-lists, detect headers, map columns, and parse employees.
    """
    header_idx, col_map = _detect_headers(rows)
    if col_map is None:
        raise ValueError(
            "Could not detect required columns. "
            "Expected columns containing: Name, Shift, Working Days/Days, Product Type/Type."
        )

    employees = []
    for i in range(header_idx + 1, len(rows)):
        row = rows[i]
        if len(row) <= max(col_map.values()):
            row = row + [""] * (max(col_map.values()) + 1 - len(row))

        name = row[col_map["name"]].strip()
        if not name or name.lower() in ("none", "nan", ""):
            continue

        shift = _parse_shift(row[col_map["shift"]]) if "shift" in col_map else None
        working_days = _parse_days(row[col_map["days"]]) if "days" in col_map else []
        content_types = _parse_types(row[col_map["type"]]) if "type" in col_map else []

        if not content_types:
            content_types = ["Content"]
        if not working_days:
            working_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

        employees.append({
            "name": name,
            "content_types": content_types,
            "working_days": working_days,
            "shift": shift,
        })

    if not employees:
        raise ValueError("No valid employee rows found in the file.")

    return employees


def _detect_headers(rows):
    """Find the header row and map column indices."""
    for idx, row in enumerate(rows):
        lower = [str(c).lower().strip() for c in row]
        col_map = {}

        for i, val in enumerate(lower):
            if not val:
                continue
            if any(k in val for k in ["name", "employee"]):
                col_map["name"] = i
            elif any(k in val for k in ["shift"]):
                col_map["shift"] = i
            elif any(k in val for k in ["day", "working", "schedule"]):
                col_map["days"] = i
            elif any(k in val for k in ["type", "product", "content", "category"]):
                col_map["type"] = i

        if "name" in col_map and len(col_map) >= 2:
            return idx, col_map

    return 0, None


def _parse_shift(val):
    val = str(val).strip().lower()
    match = re.search(r"(\d)", val)
    if match:
        s = int(match.group(1))
        if 1 <= s <= 3:
            return s
    return None


def _parse_days(val):
    val = str(val).strip()
    if not val:
        return []

    days = []
    parts = re.split(r"[,;/\s]+", val.lower())
    for p in parts:
        p = p.strip().rstrip(".")
        if p in DAY_MAP:
            full = DAY_MAP[p]
            if full not in days:
                days.append(full)

    return days


def _parse_types(val):
    val = str(val).strip()
    if not val:
        return []

    types = []
    parts = re.split(r"[,;/&]+", val.lower())
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if "content" in p or p in ("con", "cnt"):
            if "Content" not in types:
                types.append("Content")
        elif "email" in p or p in ("eml", "e-mail"):
            if "Email" not in types:
                types.append("Email")
        elif "message" in p or p in ("msg",):
            if "Message" not in types:
                types.append("Message")

    return types
