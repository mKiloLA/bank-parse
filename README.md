# Statement Parser

Statement Parser imports bank statement CSVs, helps categorize transactions, and exports an Excel workbook with spending summaries.
The desktop interface uses a dark theme, opens at a screen-fitted size, and can be resized.

## Importing statements

1. Choose one or more bank CSV files. Files do not need to be renamed or placed in a particular folder.
2. The app detects supported bank formats from their column headers and previews the transactions.
3. Review the uncategorized merchants. Select one or more groups, choose a category, and apply it once to the whole group.
4. Export the categorized transactions to Excel.

The workbook includes the summary sheets and a raw-data sheet for each submitted CSV, named for its bank. If multiple files are from the same bank, later sheets get a number suffix. The workbook filename uses the most common transaction month, such as `CombinedStatement_2026-09.xlsx`, and is saved under `~/StatementParser`.

The app recognizes the existing Capital One, Citi, Commerce, Discover, and Venmo formats, including Venmo statements with a preamble and balance summary rows. Other CSVs can use standard date, description, and amount columns, or debit and credit columns. If a file cannot be read or its columns are not recognized, the app reports the problem instead of guessing. Imported source files are not changed.

Merchant category rules are saved locally in the app's user data folder and reused on later imports. Use **Category targets** in the app to edit monthly targets.
Choose **Ignore** for transfers or credit card payments that should not count as spending. Older saved rules with the category `None` are treated as ignored.
 
## Installation (macOS)

1. Download the latest `.dmg` from the Releases page
2. Open the DMG
3. Drag **StatementParser** into the Applications folder
4. Open the app

> If macOS blocks the app: right-click → Open

The output is stored in `~/StatementParser/CombinedStatement_YYYY-MM.xlsx`.

## Installation (PC)

You can build the source with the following command:

```bash
pyinstaller --onedir --noconsole --icon=icon.icns StatementParser.py
```

The output is stored in the user's home directory under `StatementParser`, with a filename based on the most common transaction month.