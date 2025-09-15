import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import os
import sys
sys.path.append(os.path.dirname(__file__))
from ParseBankStatements import ParseBankStatements, Bank
from CategoryPopup import CategoryPopup
from GenerateOutput import GenerateOutput

class BankParseGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bank Statement Parser")
        self.geometry("1000x450")
        self.selected_folder = tk.StringVar()
        self.create_widgets()

        # Instantiate parser with GUI logger and category callback
        self.parser = ParseBankStatements(
            logger=self.log_terminal,
            get_category_callback=self._get_category_callback
        )

    def _get_category_callback(self, description, amount, date):
        popup = CategoryPopup(master=self, desc_default=description, amount=amount, date=date)
        popup.wait_window()  # Wait for user to submit
        return popup.result if popup.result else (description, "", "None")

    def create_widgets(self):
        # Top frame for folder selection and export
        top_frame = tk.Frame(self)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        folder_label = tk.Label(top_frame, text="Bank Statements Folder:")
        folder_label.pack(side=tk.LEFT)
        folder_entry = tk.Entry(top_frame, textvariable=self.selected_folder, width=50)
        folder_entry.pack(side=tk.LEFT, padx=5)
        browse_btn = tk.Button(top_frame, text="Browse", command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT, padx=5)
        export_btn = tk.Button(top_frame, text="Generate Excel File", command=self.export_data)
        export_btn.pack(side=tk.RIGHT, padx=5)

        # Main panes
        main_pane = tk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Terminal output pane
        term_frame = tk.LabelFrame(main_pane, text="Terminal Output / Input")
        self.terminal = scrolledtext.ScrolledText(term_frame, wrap=tk.WORD, height=20)
        self.terminal.pack(fill=tk.BOTH, expand=True)
        main_pane.add(term_frame, stretch="always")

        # Future functionality pane: Category Target Editor
        future_frame = tk.LabelFrame(main_pane, text="Category Targets")
        self.future_pane = tk.Frame(future_frame)
        self.future_pane.pack(fill=tk.BOTH, expand=True)
        main_pane.add(future_frame, stretch="always")

        # Add category target editor
        from CategoryTargetHandler import CategoryTargetHandler
        self.target_handler = CategoryTargetHandler()
        self.target_entries = {}
        row = 0
        tk.Label(self.future_pane, text="Category", font=("Arial", 10, "bold")).grid(row=row, column=0, padx=5, pady=2)
        tk.Label(self.future_pane, text="Target Amount", font=("Arial", 10, "bold")).grid(row=row, column=1, padx=5, pady=2)
        row += 1
        for category, amount in self.target_handler.targets.items():
            tk.Label(self.future_pane, text=category).grid(row=row, column=0, padx=5, pady=2, sticky="w")
            entry = tk.Entry(self.future_pane, width=10)
            entry.insert(0, str(amount))
            entry.grid(row=row, column=1, padx=5, pady=2)
            entry.bind("<FocusOut>", lambda e, cat=category, ent=entry: self._update_target(cat, ent))
            self.target_entries[category] = entry
            row += 1

    def _update_target(self, category, entry_widget):
        try:
            value = float(entry_widget.get())
            self.target_handler.set_target(category, value)
            self.log_terminal(f"Updated target for {category} to {value}")
        except ValueError:
            self.log_terminal(f"Invalid value for {category}: {entry_widget.get()}")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder.set(folder)
            self.log_terminal(f"Selected folder: {folder}")

    def export_data(self):
        folder = self.selected_folder.get()
        if not folder or not os.path.isdir(folder):
            self.log_terminal("Please select a valid folder containing bank statements.")
            return
        self.log_terminal(f"Parsing bank statements in: {folder}")
        try:
            bank_files = self.parser.find_bank_statements(folder)
            if not bank_files:
                self.log_terminal("No bank statement files found in the selected folder.")
                return
            statements = []
            for bank, path in bank_files.items():
                statement = self.parser.statement_from_file(bank, path)
                statements.append(statement)

            # Generate Excel output
            output_path = os.path.join("output", "expenses.xlsx")
            GenerateOutput.create_excel(statements, output_path)
            self.log_terminal(f"Parsing complete. Excel file generated at: {output_path}")
        except Exception as e:
            self.log_terminal(f"Error: {e}")

    def log_terminal(self, message):
        self.terminal.insert(tk.END, message + "\n")
        self.terminal.see(tk.END)

if __name__ == "__main__":
    app = BankParseGUI()
    app.mainloop()
