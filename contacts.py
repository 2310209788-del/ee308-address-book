import sqlite3
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HEADERS = ("Name", "Phones", "Emails", "Social accounts", "Addresses", "Favorite")


def connect(path):
    db = sqlite3.connect(path)
    db.execute("""CREATE TABLE IF NOT EXISTS contacts (
        id INTEGER PRIMARY KEY, name TEXT NOT NULL, phones TEXT NOT NULL,
        emails TEXT NOT NULL, social TEXT NOT NULL, addresses TEXT NOT NULL,
        favorite INTEGER NOT NULL DEFAULT 0)""")
    return db


def normalize(values):
    name, *details, favorite = values
    if not isinstance(name, str) or not name.strip():
        raise ValueError("Name must contain text.")
    fields = []
    for label, value in zip(HEADERS[1:5], details):
        if value is not None and not isinstance(value, str):
            raise ValueError(f"{label} must be text. Format the Excel cell as Text.")
        fields.append("\n".join(line.strip() for line in (value or "").splitlines() if line.strip()))
    flag = str(favorite).strip().lower() if favorite is not None else ""
    if flag not in ("", "0", "1", "false", "true", "no", "yes"):
        raise ValueError("Favorite must be Yes or No.")
    return (name.strip(), *fields, int(flag in ("1", "true", "yes")))


def save(db, values, contact_id=None):
    fields = normalize(values)
    with db:
        if contact_id is None:
            return db.execute("INSERT INTO contacts VALUES (NULL,?,?,?,?,?,?)", fields).lastrowid
        db.execute("UPDATE contacts SET name=?, phones=?, emails=?, social=?, addresses=?, favorite=? WHERE id=?",
                   (*fields, contact_id))
        return contact_id


def export_excel(db, path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Contacts"
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="245F58")
        sheet.column_dimensions[get_column_letter(cell.column)].width = 28
    for row in db.execute("SELECT * FROM contacts ORDER BY name, id"):
        sheet.append((*row[1:6], "Yes" if row[6] else "No"))
        for cell in sheet[sheet.max_row]:
            cell.data_type = "s"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "A2"
    book.save(path)


def import_excel(db, path):
    book = load_workbook(Path(path), read_only=True)
    try:
        rows = book.active.values
        if tuple(next(rows, ())) != HEADERS:
            raise ValueError("Expected columns: " + ", ".join(HEADERS))
        count = 0
        with db:
            for number, row in enumerate(rows, 2):
                if all(value is None for value in row):
                    continue
                try:
                    fields = normalize(row)
                except ValueError as error:
                    raise ValueError(f"Row {number}: {error}") from error
                db.execute("INSERT INTO contacts VALUES (NULL,?,?,?,?,?,?)", fields)
                count += 1
        return count
    finally:
        book.close()
