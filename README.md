# Athey Creek Book Importer v2

Version 2 enhancements:

- Automatically detects an Excel workbook located beside the EXE
- Reads exact Scripture, title, date and teaching code from each official lesson page
- Retries temporary website failures
- Skips duplicate teaching codes
- Continues through individual page errors and records warnings
- Preserves Done checkmarks and Notes
- Creates a timestamped workbook backup
- Creates a CSV import report in a `reports` folder beside the EXE
- Shows live lesson-by-lesson progress

## Build

Replace the existing repository files with the contents of this package. Committing the updated files triggers GitHub Actions. Download the artifact named `Athey-Creek-Windows-Importer-v2`.

## Use

1. Keep the workbook closed in Excel.
2. Place it beside the EXE for automatic detection, or use Browse.
3. Select the Bible book.
4. Leave Include all teachings unchecked for the Through-the-Bible path.
5. Click Import Selected Book.
6. Review the workbook and the CSV report after completion.
