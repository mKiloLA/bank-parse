# Code Structure

## Entry point

`StatementParser.py` starts `statement_parser.__main__.main()`. The package
entry point creates `BankParseGUI`, the Tkinter application window, and
coordinates the import, review, history, settings, and export workflows.

## Application modules

| Path | Responsibility |
| --- | --- |
| `statement_parser/__main__.py` | Tkinter application, page layout, user interactions, and workflow orchestration. |
| `statement_parser/source/ParseBankStatements.py` | Reads supported bank and generic CSV formats, creates transactions, and applies saved category rules. |
| `statement_parser/source/HistoryAnalysis.py` | Loads combined-statement workbooks, finds cross-file duplicates, and calculates history, category trends, and subscription candidates. |
| `statement_parser/source/GenerateOutput.py` | Builds the monthly and historical Excel workbooks. |
| `statement_parser/source/CategoryRuleStore.py` | SQLite persistence for categories, merchant rules, targets, and app settings. |
| `statement_parser/source/CategoryTargetHandler.py` | Small interface used by the GUI and workbook generator to read and update category targets. |
| `statement_parser/source/Utilities.py` | Resolves the platform-specific application data directory. |
| `statement_parser/source/data/` | Shared `Bank`, `Category`, `Statement`, and `Transaction` domain types. |
| `statement_parser/source/gui/CategoryPopup.py` | Category selection popup. |
| `statement_parser/source/gui/theme.py` | Shared colors, ttk styles, and rounded Canvas controls. |

## Main data flows

### Monthly statement import and export

1. The GUI passes selected CSV paths to `ParseBankStatements`.
2. The parser detects headers and bank format, converts rows to `Statement` and
   `Transaction` objects, and applies persisted category rules.
3. The GUI presents transactions and merchants for review; category rules and
   targets are read from or saved through `CategoryRuleStore`.
4. `GenerateOutput` creates the combined Excel workbook and its summaries.

### History analysis

1. The GUI loads prior combined workbooks through `HistoryAnalysis`.
2. Duplicate candidates are compared across source workbooks; the selected
   policy determines which occurrences are included.
3. `analyze_history` calculates monthly totals, category trends, and likely
   subscriptions for the history tables and charts.
4. `GenerateOutput` writes the separate historical workbook and duplicate
   audit.

## Persistence

The SQLite database is the source of truth for categories, merchant rules,
monthly targets, and duplicate-handling settings. `Utilities.get_appdata_path`
uses `platformdirs` to locate the per-user application data directory. Imported
CSVs and analyzed workbooks remain external source files; history analysis does
not rewrite them.

## Tests and developer tools

`tests/` contains pytest-discovered `unittest` cases:

- `test_parse_bank_statements.py`: CSV detection, parsing, and category rules.
- `test_category_rule_store.py`: SQLite persistence and settings import/export.
- `test_history_analysis.py`: workbook loading, duplicate handling, and history
  calculations.
- `test_generate_output.py`: monthly and historical workbook contents.
- `test_gui_logic.py`: GUI-facing data summaries and transaction filters.

Install development tools with `python -m pip install -r requirements-dev.txt`.
Run the unit tests with `python -m pytest -q`; Flake8, Black check, and compile
configurations are available in VS Code's Run and Debug panel.

## Refactoring notes

`BankParseGUI` currently owns most screen construction and event wiring.
Mouse-wheel direction and nested-widget routing use shared helpers, while the
history page retains its active-page guard. The statement and history
file-management dialogs still share layout but have different add/remove
callbacks; consolidating their scaffolding is a follow-up opportunity once
those page-specific flows have dedicated GUI tests. Keep CSV parsing, history
calculations, persistence, and workbook generation in their existing
non-widget modules so they remain testable without launching the GUI.
