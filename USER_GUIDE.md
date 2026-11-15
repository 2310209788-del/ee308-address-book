# Address Book User Guide

## Create and find a favorite

1. Start the application using the [Windows setup instructions](README.md#run-on-windows).
2. Select **New contact** and enter a name. To save two phone numbers for one person, put them on separate lines in **Phones**. The same rule applies to emails, social accounts, and addresses.
3. Select **Favorite**, then **Save contact**. A star appears beside the name in the contact list.
4. Select **Favorites only** to find the saved contact. To edit it, select its name, change a field, and save again. Clearing **Favorite** removes the contact from the filtered view. Clear **Favorites only** to see it again.

## Move contacts between address books

Select **Export Excel** and save the workbook as an .xlsx file. The export includes every contact even when the favorites filter is active. On the destination computer, choose **Import Excel** and open the workbook. Import adds new records; importing a workbook twice creates duplicates.

To enter records in Excel, use the exported workbook as a template. Keep the active worksheet's header row in this order:

| Name | Phones | Emails | Social accounts | Addresses | Favorite |
| --- | --- | --- | --- | --- | --- |
| Sam Lee | 0012345 | sam@example.org | | | Yes |

Format phone cells as **Text** before typing to preserve leading zeros. In Excel for Windows, use **Alt+Enter** for additional values in the same cell. Enter a name for each contact and use **Yes** or **No** for its favorite status.

If import reports a row error, correct that row and import the workbook again. The failed attempt leaves the address book as it was before the import.
