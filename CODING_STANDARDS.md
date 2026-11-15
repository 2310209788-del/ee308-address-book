# Coding Standards

## Application code

- Use descriptive names for contact fields, functions, and interface controls. Follow the naming conventions of the surrounding module.
- Keep contact storage and Excel conversion separate from interface code so that each can be tested independently.

## Excel handling

- Use the same column names for import and export. Preserve phone numbers as text so that leading zeros and international dialing prefixes survive a round trip.
- Validate required headers before changing the address book. When a value is invalid, identify the affected row and field in the error message.
- Use fictional contact details in tests. Verify that a contact with multiple phone numbers and email addresses retains every value after export and import.

## Version control

- Keep each commit focused on one change. Write a short commit message that describes the change.
- Review the diff before opening a pull request and describe how the change was verified.
