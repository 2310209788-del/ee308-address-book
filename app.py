import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from contacts import HEADERS, connect, export_excel, import_excel, save


class AddressBook:
    def __init__(self, root, db):
        self.root, self.db = root, db
        root.title("Address Book")
        root.geometry("900x590")
        root.minsize(720, 490)
        panel = ttk.Frame(root, padding=18)
        panel.pack(fill="both", expand=True)
        toolbar = ttk.Frame(panel)
        toolbar.pack(fill="x")
        ttk.Label(toolbar, text="Address Book", font=("Segoe UI", 20, "bold")).pack(side="left")
        for label, action in [("Export Excel", self.export_file), ("Import Excel", self.import_file)]:
            ttk.Button(toolbar, text=label, command=action).pack(side="right", padx=3)
        body = ttk.Frame(panel)
        body.pack(fill="both", expand=True, pady=12)
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True, padx=(0, 16))
        self.only_favorites = tk.BooleanVar()
        ttk.Checkbutton(left, text="Favorites only", variable=self.only_favorites,
                        command=self.refresh).pack(anchor="w")
        self.listing = tk.Listbox(left, width=25, exportselection=False, font=("Segoe UI", 12))
        self.listing.pack(fill="both", expand=True)
        self.listing.bind("<<ListboxSelect>>", self.select)
        ttk.Button(left, text="New contact", command=self.clear).pack(fill="x", pady=5)
        form = ttk.Frame(body)
        form.pack(side="left", fill="both", expand=True)
        ttk.Label(form, text="Contact details", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(form, text="Enter one contact method per line.").pack(anchor="w")
        self.inputs = []
        for label in HEADERS[:5]:
            ttk.Label(form, text=label).pack(anchor="w", pady=(4, 0))
            field = tk.Text(form, height=1 if label == "Name" else 2, width=32, font=("Segoe UI", 10))
            field.pack(fill="x")
            self.inputs.append(field)
        self.star = tk.BooleanVar()
        ttk.Checkbutton(form, variable=self.star, text="Favorite").pack(anchor="w", pady=5)
        ttk.Button(form, text="Save contact", command=self.submit).pack(anchor="w")
        self.status = tk.StringVar()
        ttk.Label(panel, textvariable=self.status).pack(anchor="w")
        self.refresh()

    def clear(self):
        self.contact_id = None
        self.listing.selection_clear(0, "end")
        for field in self.inputs:
            field.delete("1.0", "end")
        self.star.set(False)
        self.inputs[0].focus_set()

    def refresh(self):
        self.clear()
        self.listing.delete(0, "end")
        self.rows = list(self.db.execute(
            "SELECT * FROM contacts WHERE (?=0 OR favorite=1) ORDER BY name COLLATE NOCASE, id",
            (self.only_favorites.get(),)))
        for row in self.rows:
            self.listing.insert("end", ("★ " if row[-1] else "") + row[1])
        self.status.set(f"Contacts shown: {len(self.rows)}")

    def select(self, event=None):
        selected = self.listing.curselection()
        if selected:
            row = self.rows[selected[0]]
            self.contact_id = row[0]
            for field, value in zip(self.inputs, row[1:6]):
                field.delete("1.0", "end")
                field.insert("1.0", value)
            self.star.set(bool(row[-1]))

    def submit(self):
        values = [field.get("1.0", "end-1c") for field in self.inputs]
        try:
            save(self.db, (*values, self.star.get()), self.contact_id)
        except ValueError as error:
            messagebox.showerror("Contact details", str(error))
            return
        self.refresh()
        self.status.set("Contact saved.")

    def import_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            try:
                count = import_excel(self.db, path)
            except Exception as error:
                messagebox.showerror("Import failed", str(error))
                return
            self.refresh()
            self.status.set(f"Imported {count} contacts.")

    def export_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            try:
                export_excel(self.db, path)
            except Exception as error:
                messagebox.showerror("Export failed", str(error))
                return
            self.status.set("Address book exported.")


if __name__ == "__main__":
    folder = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "EE308AddressBook"
    folder.mkdir(parents=True, exist_ok=True)
    database = connect(folder / "contacts.sqlite3")
    try:
        window = tk.Tk()
        AddressBook(window, database)
        window.mainloop()
    finally:
        database.close()
