# Certificate Maker

This program starts a GUI that can read in bank statements and outputs an Excel sheet with the data.
 
## Installation (macOS)

1. Download the latest `.dmg` from the Releases page
2. Open the DMG
3. Drag **StatementParser** into the Applications folder
4. Open the app

> If macOS blocks the app: right-click → Open

The output will be stored in `~/Documents/StatementParser/combinedStatements.xlsx`.

## Installation (PC)

You can build the source with the following command:

```bash
pyinstaller --onedir --noconsole --icon=icon.icns StatementParser.py
```

After running on PC, the output will be stored in file called `output/StatementParser.xlsx` in the folder containing the app.