# Statement Parser

Statement Parser imports bank statement CSVs, helps categorize transactions, and exports an Excel workbook with spending summaries.
The desktop interface uses a dark theme, opens at a screen-fitted size, and can be resized.

## Importing statements

1. Choose bank CSV files, or drag and drop them anywhere in the app window. New files are added to the current import; use **Manage Files** to add more or remove selected files. Removing a file only removes it from the current import and does not delete the original.
2. The app detects supported bank formats from their column headers and previews the transactions.
3. Search and filter the transaction preview by merchant, category, source, or month.
4. Review the uncategorized merchants and their source bank. Select one or more groups, choose a category, and apply it once to the whole group. Enter shorter **Rule Match Text** (for example, `google`) to match similar descriptions; blank uses the full transaction description. Choose **Contains** or **Exact**.
5. Use **Monthly Overview** to compare spending and income with saved category targets.
6. Review potential duplicate transactions when prompted. Matches require the same date, amount, and place across different files.
7. Export the categorized transactions to Excel.

The workbook opens to a dashboard with income, spending, net cash flow, savings rate, category budgets, top merchants, and monthly and daily spending charts. Monthly targets are scaled to the number of months covered by the imported transactions. The workbook also includes category and income summaries, detailed purchase and income tables, and a raw-data sheet for each submitted CSV, named for its bank. If multiple files are from the same bank, later sheets get a number suffix. The workbook filename uses the most common transaction month, such as `CombinedStatement_2026-09.xlsx`, and is saved under `~/StatementParser`.

## History and trends

Use **History & Trends** to add one or more previously exported `CombinedStatement_YYYY-MM.xlsx` workbooks. The app lists their date ranges, detects cross-workbook duplicate candidates, and analyzes the consolidated history without changing the source workbooks. Review or remove selected duplicate occurrences before analysis, or choose a duplicate-handling policy in **Settings**:

- **Warn and review** keeps all matches unless you exclude them; the app asks whether to review before analysis or export.
- **Automatically exclude duplicates** keeps the earliest copy across the selected workbooks.
- **Keep all transactions** disables duplicate filtering.

The category table reports average monthly spending and estimated monthly change, calculated with a least-squares trend across the selected period (including months with no spending). Likely subscriptions are suggestions for the same merchant and exact amount appearing in at least three distinct months. Export creates a separate historical workbook with an overview, category trends, subscription candidates, transaction details, and a duplicate audit. The analysis processes transactions in a small number of passes and does not repeatedly scan the source workbooks.

The app recognizes the existing Capital One, Citi, Commerce, Discover, and Venmo formats, including Venmo statements with a preamble and balance summary rows. Other CSVs can use standard date, description, and amount columns, or debit and credit columns. If a file cannot be read or its columns are not recognized, the app reports the problem instead of guessing. Imported source files are not changed.

Categories, merchant rules, and monthly targets are saved in a local SQLite database. Review rules apply to matching transactions already loaded as well as future imports. Use **Category targets** to add, rename, and remove categories and edit their monthly targets. Custom categories are available throughout categorization, rules, and reports. Removing a category also removes its saved rules. Use **Category rules** to search, sort by clicking a column heading, review, add, edit, or delete rules; each rule can match an exact description or any description containing its match text. Choose **Ignore** for transfers or credit card payments that should not count as spending.

The database is stored in the platform's user data directory: `~/Library/Application Support/StatementParser` on macOS, `~/.local/share/StatementParser` on Linux, and the per-user application data directory on Windows. SQLite is the source of truth for saved rules and targets.

Use **Settings** to choose duplicate handling, save a portable `StatementParser_Settings.db` file beside the combined statements under `~/StatementParser`, or import a settings file from another computer. Import replaces the current categories, rules, targets, and duplicate-handling policy after first saving a timestamped backup there.

## Run from source

Install the requirements and launch the app:

```bash
python3 -m pip install -r requirements.txt
python3 StatementParser.py
```

On Ubuntu, install Tkinter if it is not already available:

```bash
sudo apt install python3-tk
```

## Installation (macOS)

1. Download the latest `.dmg` from the Releases page
2. Open the DMG
3. Drag **StatementParser** into the Applications folder
4. Open the app

> If macOS blocks the app: right-click → Open

The output is stored in `~/StatementParser/CombinedStatement_YYYY-MM.xlsx`.

## Build from source

PyInstaller builds are specific to the operating system they run on. Build the macOS app on macOS:

```bash
python3 -m pip install -r requirements.txt
pyinstaller StatementParser.spec
```

Build an Ubuntu/Linux distribution on Ubuntu:

```bash
python3 -m pip install -r requirements.txt
pyinstaller --onedir --noconsole --icon=icon.png StatementParser.py
```

The generated application is in `dist/StatementParser`. Exported workbooks are stored in the user's home directory under `StatementParser`, with a filename based on the most common transaction month.