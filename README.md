# Athey Creek Book Importer — Final Version 3

This final version is designed for the book-by-book workflow.

## Features

- Select any of the 66 Bible books
- Preview every matching lesson before Excel is changed
- Reads exact Scripture, title, date, teaching code and lesson URL from each official lesson page
- Imports only Through the Bible lessons unless **Include all teachings** is checked
- Verify the selected book already in the workbook against the current official pages
- Preserves Done checkmarks and Notes by teaching code
- Replaces only the selected book
- Keeps Bible-book order
- Creates a timestamped workbook backup
- Creates a CSV import report
- Detects duplicate teaching codes
- Retries temporary website failures
- Automatically suggests an `.xlsx` workbook placed beside the EXE

## Build in the existing GitHub repository

Upload and replace the visible files in this package. Then replace the workflow file at:

`.github/workflows/build-windows-exe.yml`

GitHub Actions will create an artifact named:

`Athey-Creek-Windows-Importer-Final`

## Normal use

1. Close the workbook in Excel.
2. Open `Athey_Creek_Book_Importer.exe`.
3. Select the next Bible book.
4. Select the workbook.
5. Click **1. Preview Selected Book**.
6. Review the exact Scripture, title and date in the preview table.
7. Click **2. Import Previewed Book**.
8. Reopen the workbook after the completion message.

Use **Verify Workbook Book** whenever you want to compare an already imported book to the live official pages without changing Excel.
