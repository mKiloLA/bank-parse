# Certificate Maker

This program starts a GUI that can read in bank statements and outputs an Excel sheet with the data.

## Installation

You can build the source with the following command:

```bash
pyinstaller --onefile --noconsole StatementParser.py
```

After running, the output will be stored in file called `output/expenses.xlsx`.
