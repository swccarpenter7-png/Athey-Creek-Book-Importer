# Athey Creek Book Importer

A standalone Windows application that imports one Bible book at a time into the existing two-tab Excel workbook.

## What it imports

For every selected lesson, the program reads the **individual official Athey Creek lesson page** and records:

- Bible book
- Exact Scripture reference shown under **Book**
- Exact lesson title
- Exact lesson date
- Teaching code
- Clickable **Listen** link
- Official lesson URL

By default, it imports only lesson titles beginning with **Through the Bible**. A checkbox can include every teaching listed for the selected Bible book.

## Workbook requirements

The `.xlsx` file must contain a worksheet named:

- `Master Index`

A `Dashboard` worksheet is updated or created automatically. The intended workbook contains only these two tabs.

The program automatically creates a timestamped backup beside the workbook before making changes.

## Build the EXE using GitHub Actions

1. Create a free GitHub account or sign in.
2. Create a new repository. A private repository is fine.
3. Upload **all files and folders from this project**, including the hidden `.github` folder.
4. Commit the files to the `main` branch.
5. Open the repository's **Actions** tab.
6. Select **Build Windows EXE**.
7. Click **Run workflow**, then click the green **Run workflow** button.
8. Wait for the build to finish with a green checkmark.
9. Open the completed workflow run.
10. Under **Artifacts**, download **Athey-Creek-Windows-Importer**.
11. Unzip the downloaded package.
12. Double-click `Athey_Creek_Book_Importer.exe`.

No Python installation is needed to run the compiled EXE.

## Use the importer

1. Close your Excel workbook.
2. Open `Athey_Creek_Book_Importer.exe`.
3. Select the Bible book.
4. Browse to the workbook.
5. Leave **Include all teachings** unchecked for the Through-the-Bible listening path.
6. Click **Import Selected Book**.
7. Wait for the completion message, then reopen the workbook.

## Updating a book

Running the same book again replaces that book's imported rows with current official page data. It preserves existing Done checkmarks and Notes when the teaching code still matches.

## Windows security message

Because this personal application is not digitally code-signed, Windows SmartScreen may display **Windows protected your PC** the first time it runs. Confirm that you built/downloaded it from your own GitHub repository, then choose **More info** and **Run anyway**.

## Important limitations

- Internet access is required while importing.
- The Athey Creek website is the source of truth.
- A website redesign may eventually require parser updates.
- Keep your workbook closed in Excel during import.
