import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from contacts import HEADERS, connect, export_excel, import_excel, save


class ExcelExchangeTests(unittest.TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / "contacts.xlsx"
        self.source = connect(":memory:")
        self.destination = connect(":memory:")
        self.addCleanup(self.source.close)
        self.addCleanup(self.destination.close)

    def test_round_trip_preserves_contact_details_and_favorites(self):
        contacts = [
            ("Alex", "00123\n+44001", "a@example.org\nb@example.org", "=alex", "North\nSouth", 1),
            ("Sam", "", "", "", "", 0),
        ]
        for contact in contacts:
            save(self.source, contact)
        export_excel(self.source, self.path)
        book = load_workbook(self.path)
        try:
            sheet = book.active
            self.assertEqual(tuple(cell.value for cell in sheet[1]), HEADERS)
            self.assertEqual(sheet.max_row, 3)
            self.assertEqual(sheet["D2"].data_type, "s")
        finally:
            book.close()
        self.assertEqual(import_excel(self.destination, self.path), 2)
        rows = self.destination.execute(
            "SELECT name, phones, emails, social, addresses, favorite FROM contacts ORDER BY name"
        ).fetchall()
        self.assertEqual(rows, contacts)

    def test_invalid_row_rolls_back_only_the_attempted_import(self):
        save(self.destination, ("Existing", "", "", "", "", 0))
        save(self.source, ("Valid", "", "", "", "", 0))
        export_excel(self.source, self.path)
        book = load_workbook(self.path)
        try:
            book.active.append(("Invalid", 123, "", "", "", "No"))
            book.save(self.path)
        finally:
            book.close()
        with self.assertRaisesRegex(ValueError, "Row 3: Phones"):
            import_excel(self.destination, self.path)
        self.assertEqual(self.destination.execute("SELECT name FROM contacts").fetchall(), [("Existing",)])
