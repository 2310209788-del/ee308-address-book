# EE308FZ Address Book

This pair programming project for EE308FZ Software Engineering focuses on organizing contacts and exchanging address book data through Excel.

## Feature acceptance guide

Use the [project directory build](https://github.com/2310209788-del/ee308-address-book/tree/382f3c8b0c8f5ad8f6959b2d688253affbb23f08) for all checks below. It includes the role and search features added after the first release candidate.

### Course requirements

| Grading item | Points | How to verify |
| --- | ---: | --- |
| Bookmark contacts | 25 | Save Alex as a favorite and Sam as a regular contact. Enable **Favorites only**. Alex has a star and remains visible. Clear the filter to see both contacts. |
| Multiple contact methods | 10 | Give Alex two phone numbers and two email addresses, one per line. Save and reopen Alex. Both numbers and both addresses are present. |
| Excel export and import | 25 | Export both contacts. Check that the spreadsheet has labeled columns and one row per person. Change one name in a copy and import it. Open the newly imported contact and verify its details. |
| Interface style and usability | 10 | Select a contact in the left pane, edit it in the right pane, and save. Check that labels, favorite status, and the selected contact are easy to identify. |

### Features added by software_clan

| Addition | How to verify |
| --- | --- |
| Project roles and two-line contact entries | Set Alex's **Project role** to **Testing**, or type a custom role. Save and check that the role appears beneath the name. |
| Live name or role search | Enter part of a name, then try `test` to find the Testing role. Enable **Favorites only** while searching to apply both filters together. The teal highlight identifies the selected contact. |
| Excel role compatibility | Export and reimport a contact with a role. The role is preserved. Import an older six-column workbook without **Role**. Its contacts are added with an empty role. |

The course lists public cloud deployment as a separate optional bonus (up to 15 points). This submission runs locally on Windows and claims no deployment bonus.

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
