# EE308FZ Address Book

This pair programming project for EE308FZ Software Engineering focuses on organizing contacts and exchanging address book data through Excel.

## Application requirements

- Mark contacts as favorites and make them easy to find in the contact list.
- Store multiple contact details for one person, including phone numbers, email addresses, social media accounts, and postal addresses.
- Export one row per contact with labeled columns. Import workbooks in the same format and preserve all contact details.

## Find a project teammate

Our campus teams often know who they need by responsibility before they know the person's name. Assign an optional project role and search by name or responsibility. Combine a search such as Testing with Favorites only to find the right contact before a demo. Existing contacts and older six-column workbooks remain compatible.

## Responsibilities

The repository owner leads application development and integration. The teammate leads project documentation.

[Assignment brief](https://bbs.csdn.net/topics/620593436) | [Coding standards](CODING_STANDARDS.md)

## Run on Windows

Install Python 3.12 with Tcl/Tk support. From the project directory run:

```powershell
py -3.12 -m pip install -r requirements.txt
py -3.12 app.py
```

Choose **New contact**, enter a name, and put each phone number, email, social account, or address on its own line. Select **Favorite** before saving to include the contact in the favorites view. Select an existing contact to edit it. The application stores contacts in a per-user SQLite database under `%LOCALAPPDATA%\EE308AddressBook`.

**Export Excel** writes the full address book, including contacts hidden by the favorites filter. To create an import template, export an empty address book. Keep the column names and order, enter phone numbers as text, and use line breaks inside a cell for multiple values. **Import Excel** appends all valid rows as new contacts. Reimporting the same workbook creates additional copies. If a row fails validation, the entire import is rolled back.

See the [user guide](USER_GUIDE.md) for a worked example. Run the Excel exchange tests with `py -3.12 -m unittest -v test_excel`.
