import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.append(os.path.dirname(__file__))

from .source.CategoryTargetHandler import CategoryTargetHandler
from .source.GenerateOutput import GenerateOutput
from .source.ParseBankStatements import ParseBankStatements
from .source.data.DataTypes import Category
from .source.data.Statement import Statement
from .source.data.Transaction import Transaction


class BankParseGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Statement Parser")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = max(800, min(1500, screen_width - 40))
        height = max(600, min(960, screen_height - 80))
        self.geometry(
            f"{width}x{height}+{(screen_width - width) // 2}+{(screen_height - height) // 2}"
        )
        self.minsize(min(1040, width), min(700, height))
        self.resizable(True, True)
        self.configure(background="#10151f")

        self.parser = ParseBankStatements(logger=self.log_status)
        self.target_handler = CategoryTargetHandler()
        self.statements: list[Statement] = []
        self.review_groups: dict[str, list[Transaction]] = {}
        self.selected_files: list[str] = []
        self.category_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.file_summary_var = tk.StringVar(value="No statements loaded")
        self.summary_var = tk.StringVar(value="Choose statement CSV files to get started.")
        self.status_var = tk.StringVar(value="Ready")

        self._configure_styles()
        self._create_widgets()

    def _configure_styles(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            ".", background="#10151f", foreground="#e7edf5", font=("Arial", 16)
        )
        style.configure("App.TFrame", background="#10151f")
        style.configure("Card.TFrame", background="#1a2230")
        style.configure(
            "Title.TLabel",
            background="#10151f",
            foreground="#f2f5fa",
            font=("Arial", 34, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#10151f",
            foreground="#a4afbf",
            font=("Arial", 14),
        )
        style.configure(
            "CardTitle.TLabel",
            background="#1a2230",
            foreground="#edf2f8",
            font=("Arial", 18, "bold"),
        )
        style.configure(
            "Muted.TLabel",
            background="#1a2230",
            foreground="#a4afbf",
            font=("Arial", 14),
        )
        style.configure(
            "Primary.TButton",
            background="#526de5",
            foreground="#ffffff",
            font=("Arial", 16),
            padding=(16, 11),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("pressed", "#4059c3"),
                ("active", "#627cf0"),
                ("disabled", "#34415a"),
            ],
            foreground=[("disabled", "#9aa6ba")],
        )
        style.configure(
            "TButton",
            background="#263246",
            foreground="#e7edf5",
            font=("Arial", 15),
            padding=(15, 11),
            bordercolor="#3a475c",
            focuscolor="#526de5",
        )
        style.map(
            "TButton",
            background=[("pressed", "#303e55"), ("active", "#303e55")],
            foreground=[("disabled", "#79869a")],
        )
        style.configure(
            "TEntry",
            fieldbackground="#131b29",
            foreground="#e7edf5",
            font=("Arial", 15),
            padding=10,
        )
        style.map(
            "TEntry",
            fieldbackground=[("disabled", "#202a39")],
            foreground=[("disabled", "#8d9aae")],
        )
        style.configure(
            "TCombobox",
            fieldbackground="#131b29",
            background="#263246",
            foreground="#e7edf5",
            arrowcolor="#c5cfdd",
            font=("Arial", 15),
            padding=9,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#131b29"), ("disabled", "#202a39")],
            foreground=[("readonly", "#e7edf5"), ("disabled", "#8d9aae")],
            selectbackground=[("readonly", "#354866")],
            selectforeground=[("readonly", "#ffffff")],
        )
        style.configure(
            "Treeview",
            rowheight=44,
            background="#151d2a",
            fieldbackground="#151d2a",
            foreground="#e0e7f0",
            font=("Arial", 14),
            bordercolor="#303b4d",
            lightcolor="#303b4d",
            darkcolor="#303b4d",
        )
        style.configure(
            "Treeview.Heading",
            background="#263246",
            foreground="#cbd5e2",
            font=("Arial", 13, "bold"),
            padding=(10, 13),
        )
        style.map(
            "Treeview",
            background=[("selected", "#354d78")],
            foreground=[("selected", "#ffffff")],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#303e55")],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#263246",
            troughcolor="#151d2a",
            arrowcolor="#c5cfdd",
            bordercolor="#151d2a",
            arrowsize=18,
        )
        style.configure("TPanedwindow", background="#10151f")

    def _create_widgets(self):
        shell = ttk.Frame(self, style="App.TFrame", padding=(28, 24, 28, 20))
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(2, weight=1)

        header = ttk.Frame(shell, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Statement Parser", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Import your statements, review new merchants, and export your spending.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        toolbar = ttk.Frame(shell, style="App.TFrame")
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        toolbar.columnconfigure(1, weight=1)
        ttk.Button(
            toolbar,
            text="Choose CSV files",
            style="Primary.TButton",
            command=self.choose_files,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            toolbar, textvariable=self.file_summary_var, style="Subtitle.TLabel"
        ).grid(row=0, column=1, sticky="w", padx=(14, 8))
        ttk.Button(
            toolbar, text="Category targets", command=self.edit_targets
        ).grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.export_button = ttk.Button(
            toolbar,
            text="Export Excel",
            command=self.export_data,
            state=tk.DISABLED,
        )
        self.export_button.grid(row=0, column=3, sticky="e")

        content = ttk.Panedwindow(shell, orient=tk.HORIZONTAL)
        content.grid(row=2, column=0, sticky="nsew")
        transaction_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        review_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        content.add(transaction_card, weight=7)
        content.add(review_card, weight=4)

        transaction_card.columnconfigure(0, weight=1)
        transaction_card.rowconfigure(2, weight=1)
        ttk.Label(
            transaction_card, text="Transaction preview", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            transaction_card, textvariable=self.summary_var, style="Muted.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        transaction_columns = ("date", "merchant", "amount", "category", "bank")
        self.transaction_tree = ttk.Treeview(
            transaction_card,
            columns=transaction_columns,
            show="headings",
            selectmode="browse",
        )
        for column, heading, width, anchor in (
            ("date", "Date", 120, "w"),
            ("merchant", "Merchant / description", 300, "w"),
            ("amount", "Amount", 120, "e"),
            ("category", "Category", 160, "w"),
            ("bank", "Source", 130, "w"),
        ):
            self.transaction_tree.heading(column, text=heading)
            self.transaction_tree.column(
                column, width=width, minwidth=65, anchor=anchor
            )
        transaction_scroll = ttk.Scrollbar(
            transaction_card, orient=tk.VERTICAL, command=self.transaction_tree.yview
        )
        self.transaction_tree.configure(yscrollcommand=transaction_scroll.set)
        self.transaction_tree.grid(row=2, column=0, sticky="nsew")
        transaction_scroll.grid(row=2, column=1, sticky="ns")

        review_card.columnconfigure(0, weight=1)
        review_card.rowconfigure(2, weight=1)
        ttk.Label(
            review_card, text="Review merchants", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            review_card,
            text="Choose a category, or Ignore transfers and payments you don't want counted.",
            style="Muted.TLabel",
            wraplength=310,
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        group_columns = ("merchant", "count", "total")
        self.review_tree = ttk.Treeview(
            review_card,
            columns=group_columns,
            show="headings",
            selectmode="extended",
            height=12,
        )
        self.review_tree.heading("merchant", text="Merchant")
        self.review_tree.heading("count", text="#", anchor="center")
        self.review_tree.heading("total", text="Total", anchor="e")
        self.review_tree.column("merchant", width=230, minwidth=110)
        self.review_tree.column("count", width=55, minwidth=42, anchor="center")
        self.review_tree.column("total", width=115, minwidth=80, anchor="e")
        review_scroll = ttk.Scrollbar(
            review_card, orient=tk.VERTICAL, command=self.review_tree.yview
        )
        self.review_tree.configure(yscrollcommand=review_scroll.set)
        self.review_tree.grid(row=2, column=0, sticky="nsew")
        review_scroll.grid(row=2, column=1, sticky="ns")
        self.review_tree.bind("<<TreeviewSelect>>", self._select_review_group)

        ttk.Label(
            review_card, text="Category", style="CardTitle.TLabel"
        ).grid(row=3, column=0, sticky="w", pady=(16, 5))
        self.category_combo = ttk.Combobox(
            review_card,
            textvariable=self.category_var,
            values=[category.value for category in Category if category != Category.NONE],
            state="readonly",
        )
        self.category_combo.grid(row=4, column=0, sticky="ew")
        ttk.Label(
            review_card, text="Display title (optional)", style="CardTitle.TLabel"
        ).grid(row=5, column=0, sticky="w", pady=(12, 5))
        ttk.Entry(review_card, textvariable=self.title_var).grid(
            row=6, column=0, sticky="ew"
        )
        ttk.Button(
            review_card,
            text="Apply to selected merchants",
            style="Primary.TButton",
            command=self.apply_categories,
        ).grid(row=7, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(
            review_card,
            text="Use Ctrl-click (or Command-click on Mac) to select several.",
            style="Muted.TLabel",
            wraplength=310,
        ).grid(row=8, column=0, sticky="w", pady=(8, 0))

        status_bar = ttk.Frame(shell, style="App.TFrame")
        status_bar.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(status_bar, textvariable=self.status_var, style="Subtitle.TLabel").pack(
            side=tk.LEFT
        )

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Choose bank statement CSV files",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
        )
        if not paths:
            return

        statements = []
        imported_paths = []
        errors = []
        for path in paths:
            try:
                statements.append(self.parser.statement_from_file(None, path))
                imported_paths.append(path)
            except (OSError, ValueError) as error:
                errors.append(f"{os.path.basename(path)}: {error}")

        if not statements:
            messagebox.showerror(
                "Could not import statements",
                "\n\n".join(errors) or "No transactions could be read from those files.",
                parent=self,
            )
            return

        self.statements = statements
        self.selected_files = imported_paths
        self.file_summary_var.set(
            f"{len(imported_paths)} CSV file{'s' if len(imported_paths) != 1 else ''} loaded"
        )
        self.export_button.configure(state=tk.NORMAL)
        self._refresh_preview()

        if errors:
            messagebox.showwarning(
                "Some files were skipped",
                "\n\n".join(errors),
                parent=self,
            )

    def _refresh_preview(self):
        self.review_groups = {}
        for item in self.transaction_tree.get_children():
            self.transaction_tree.delete(item)
        for item in self.review_tree.get_children():
            self.review_tree.delete(item)

        transaction_count = 0
        for statement_index, statement in enumerate(self.statements):
            for transaction_index, transaction in enumerate(statement.transactions):
                transaction_count += 1
                category = transaction.category
                category_label = (
                    "Needs review" if category == Category.NONE else category.value
                )
                self.transaction_tree.insert(
                    "",
                    tk.END,
                    iid=f"transaction-{statement_index}-{transaction_index}",
                    values=(
                        transaction.date.strftime("%Y-%m-%d"),
                        transaction.description,
                        self._format_amount(transaction.amount),
                        category_label,
                        statement.bank.value,
                    ),
                )
                if category == Category.NONE:
                    key = self.parser.normalize_description(transaction.description)
                    self.review_groups.setdefault(key, []).append(transaction)

        for index, (key, transactions) in enumerate(
            sorted(
                self.review_groups.items(),
                key=lambda item: (-len(item[1]), item[0]),
            )
        ):
            total = sum(transaction.amount for transaction in transactions)
            label = transactions[0].description
            self.review_tree.insert(
                "",
                tk.END,
                iid=f"merchant-{index}",
                values=(label, len(transactions), self._format_amount(total)),
            )
            self.review_tree.set(f"merchant-{index}", "merchant", label)
            self.review_tree.item(
                f"merchant-{index}", tags=(key,)
            )

        needs_review = sum(len(transactions) for transactions in self.review_groups.values())
        self.export_button.configure(
            state=tk.NORMAL if transaction_count and not needs_review else tk.DISABLED
        )
        self.summary_var.set(
            f"{transaction_count} transaction{'s' if transaction_count != 1 else ''} "
            f"- {needs_review} need categorizing"
        )
        self.status_var.set(
            "Categorize every merchant group before exporting."
            if needs_review
            else "All transactions are categorized. You can export your spreadsheet."
            if transaction_count
            else "No transactions found in the selected files."
        )

    def _selected_group_keys(self) -> list[str]:
        keys = []
        for item in self.review_tree.selection():
            tags = self.review_tree.item(item, "tags")
            if tags:
                keys.append(tags[0])
        return keys

    def _select_review_group(self, _event=None):
        selected_keys = self._selected_group_keys()
        if len(selected_keys) == 1:
            transaction = self.review_groups[selected_keys[0]][0]
            self.title_var.set(transaction.description)
        else:
            self.title_var.set("")
        self.category_var.set("")

    def apply_categories(self):
        keys = self._selected_group_keys()
        category = next(
            (
                item
                for item in Category
                if item.value == self.category_var.get() and item != Category.NONE
            ),
            None,
        )
        if not keys:
            messagebox.showinfo(
                "Select merchants", "Select one or more merchants to categorize.", parent=self
            )
            return
        if category is None:
            messagebox.showinfo(
                "Choose a category", "Choose a category before applying it.", parent=self
            )
            return

        title = self.title_var.get().strip()
        rules = {}
        for key in keys:
            transactions = self.review_groups[key]
            display_title = title or transactions[0].description
            rules[transactions[0].description] = (display_title, category)
        try:
            self.parser.save_category_rules(rules)
        except OSError as error:
            messagebox.showerror(
                "Could not save merchant rules",
                f"Your categories were not applied because the rules could not be saved:\n{error}",
                parent=self,
            )
            return
        except ValueError as error:
            messagebox.showerror("Could not categorize merchants", str(error), parent=self)
            return

        for key in keys:
            transactions = self.review_groups[key]
            display_title = title or transactions[0].description
            for transaction in transactions:
                transaction.title = display_title
                transaction.category = category
        self._refresh_preview()
        self.status_var.set(
            f"Categorized {len(keys)} merchant group{'s' if len(keys) != 1 else ''}. "
            "The rules will be remembered for future imports."
        )

    def export_data(self):
        if not self.statements:
            messagebox.showinfo(
                "No statements loaded", "Choose statement CSV files before exporting.", parent=self
            )
            return
        if any(
            transaction.category == Category.NONE
            for statement in self.statements
            for transaction in statement.transactions
        ):
            messagebox.showinfo(
                "Review categories first",
                "Categorize the remaining merchant groups before exporting so no transactions are left out.",
                parent=self,
            )
            return
        output_path = GenerateOutput.default_output_path(self.statements)
        source_files = [
            (statement.bank, path)
            for statement, path in zip(self.statements, self.selected_files)
        ]
        try:
            GenerateOutput.create_excel(
                self.statements, output_path, source_files=source_files
            )
        except (OSError, ValueError) as error:
            messagebox.showerror(
                "Export failed", f"Could not create the Excel file:\n{error}", parent=self
            )
            return
        self.status_var.set(f"Excel file created: {output_path}")
        messagebox.showinfo(
            "Export complete", f"Your spreadsheet is ready:\n{output_path}", parent=self
        )

    def edit_targets(self):
        dialog = tk.Toplevel(self)
        dialog.configure(background="#10151f")
        dialog.title("Monthly category targets")
        dialog.transient(self)
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Monthly category targets", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )

        entries = {}
        for row, (category, amount) in enumerate(
            self.target_handler.targets.items(), start=1
        ):
            ttk.Label(frame, text=category).grid(
                row=row, column=0, sticky="w", padx=(0, 18), pady=4
            )
            entry = ttk.Entry(frame, width=14)
            entry.insert(0, str(amount))
            entry.grid(row=row, column=1, sticky="e", pady=4)
            entries[category] = entry

        def save_targets():
            try:
                updated = {
                    category: float(entry.get())
                    for category, entry in entries.items()
                }
            except ValueError:
                messagebox.showerror(
                    "Invalid target",
                    "Enter a valid number for every category target.",
                    parent=dialog,
                )
                return
            try:
                for category, amount in updated.items():
                    self.target_handler.set_target(category, amount)
            except OSError as error:
                messagebox.showerror(
                    "Could not save targets", str(error), parent=dialog
                )
                return
            dialog.destroy()
            self.status_var.set("Monthly category targets saved.")

        buttons = ttk.Frame(frame)
        buttons.grid(row=len(entries) + 1, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        ttk.Button(
            buttons, text="Save targets", style="Primary.TButton", command=save_targets
        ).pack(side=tk.RIGHT)

    @staticmethod
    def _format_amount(amount: float) -> str:
        return f"-${abs(amount):,.2f}" if amount < 0 else f"${amount:,.2f}"

    def log_status(self, message: str):
        self.status_var.set(message)


def main():
    app = BankParseGUI()
    app.mainloop()
