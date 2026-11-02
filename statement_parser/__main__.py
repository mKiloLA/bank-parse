import colorsys
import os
import sqlite3
import sys
import tkinter as tk
from zipfile import BadZipFile
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from openpyxl.utils.exceptions import InvalidFileException
from tkinterdnd2 import COPY, DND_FILES, TkinterDnD

sys.path.append(os.path.dirname(__file__))

from .source.CategoryTargetHandler import CategoryTargetHandler
from .source.GenerateOutput import GenerateOutput
from .source.HistoryAnalysis import (
    analyze_history,
    duplicate_occurrences_to_skip,
    filter_duplicate_occurrences,
    find_duplicate_groups,
    load_combined_statement,
)
from .source.ParseBankStatements import ParseBankStatements
from .source.data.DataTypes import Category, category_from_name, category_name
from .source.data.Statement import Statement
from .source.data.Transaction import Transaction
from .source.gui.theme import (
    ACCENT,
    BACKGROUND,
    BORDER,
    LOG_BACKGROUND,
    METRIC_BUDGET,
    METRIC_SPENT,
    RoundedButton,
    RoundedCard,
    STATUS_OVER,
    STATUS_UNDER,
    SURFACE,
    SURFACE_RAISED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    configure_styles,
)


class BankParseGUI(TkinterDnD.Tk):
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
        self.configure(background=BACKGROUND)

        self.parser = ParseBankStatements(logger=self.log_status)
        self.target_handler = CategoryTargetHandler()
        self.statements: list[Statement] = []
        self.review_groups: dict[str, list[Transaction]] = {}
        self.selected_files: list[str] = []
        self.monthly_duplicate_exclusions: set[tuple[int, int]] = set()
        self.monthly_duplicate_groups = []
        self.monthly_effective_exclusions: set[tuple[int, int]] = set()
        self.monthly_analysis_statements: list[Statement] = []
        self.monthly_duplicates_reviewed = False
        self.history_files: list[str] = []
        self.history_statements: list[Statement] = []
        self.history_duplicate_exclusions: set[tuple[int, int]] = set()
        self.history_analysis = None
        self.history_duplicate_groups = []
        self.history_effective_exclusions: set[tuple[int, int]] = set()
        self.history_analysis_statements: list[Statement] = []
        self.history_duplicates_reviewed = False
        self.history_category_visibility = {}
        self._category_rules_refresh = None
        self._category_targets_refresh = None
        self.category_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.rule_pattern_var = tk.StringVar()
        self.rule_match_type_var = tk.StringVar(value="Contains")
        self.file_summary_var = tk.StringVar(value="Manage Files (0)")
        self.header_subtitle_var = tk.StringVar(
            value="Choose or drop CSV files, review new merchants, and export your spending."
        )
        self.summary_var = tk.StringVar(value="Choose statement CSV files to get started.")
        self.status_var = tk.StringVar(value="Ready")
        self.transaction_search_var = tk.StringVar()
        self.transaction_category_filter_var = tk.StringVar(value="All Categories")
        self.transaction_source_filter_var = tk.StringVar(value="All Sources")
        self.transaction_month_filter_var = tk.StringVar(value="All Months")
        self.overview_month_var = tk.StringVar()
        self.duplicate_policy_var = tk.StringVar()
        self.history_summary_var = tk.StringVar(
            value="Add CombinedStatement workbooks to analyze multiple months."
        )

        self._configure_styles()
        self._create_widgets()
        self._register_drop_targets(self)

    def _register_drop_targets(self, widget):
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", self._handle_file_drop)
        for child in widget.winfo_children():
            self._register_drop_targets(child)

    def _configure_styles(self):
        configure_styles(ttk.Style(self))

    def _create_widgets(self):
        shell = ttk.Frame(self, style="App.TFrame", padding=(28, 24, 28, 20))
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(3, weight=1)

        header = ttk.Frame(shell, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Statement Parser", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            textvariable=self.header_subtitle_var,
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        toolbar = ttk.Frame(shell, style="App.TFrame")
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        toolbar.columnconfigure(1, weight=1)
        self.choose_files_button = RoundedButton(
            toolbar,
            text="Choose CSV Files",
            style="Primary.TButton",
            command=self.choose_files,
        )
        self.choose_files_button.grid(row=0, column=0, sticky="w")
        self.manage_files_button = RoundedButton(
            toolbar,
            textvariable=self.file_summary_var,
            command=self.manage_files,
        )
        self.manage_files_button.grid(row=0, column=1, sticky="w", padx=(14, 8))
        self.category_targets_button = RoundedButton(
            toolbar, text="Category Targets", command=self.edit_targets
        )
        self.category_targets_button.grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.category_rules_button = RoundedButton(
            toolbar, text="Category Rules", command=self.edit_category_rules
        )
        self.category_rules_button.grid(row=0, column=3, sticky="e", padx=(0, 8))
        self.settings_button = RoundedButton(
            toolbar, text="Settings", command=self.open_settings
        )
        self.settings_button.grid(row=0, column=4, sticky="e", padx=(0, 8))
        self.export_button = RoundedButton(
            toolbar,
            text="Export Excel",
            command=self.export_data,
            state=tk.DISABLED,
        )
        self.export_button.grid(row=0, column=5, sticky="e")

        tab_bar = ttk.Frame(shell, style="App.TFrame")
        tab_bar.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        for column in range(3):
            tab_bar.columnconfigure(column, weight=1, uniform="workspace-tabs")

        self.workspace = ttk.Frame(shell, style="App.TFrame")
        self.workspace.grid(row=3, column=0, sticky="nsew")
        self.workspace.columnconfigure(0, weight=1)
        self.workspace.rowconfigure(0, weight=1)
        review_page = ttk.Frame(self.workspace, style="App.TFrame")
        overview_page = ttk.Frame(self.workspace, style="App.TFrame", padding=20)
        self.history_page = ttk.Frame(
            self.workspace, style="App.TFrame"
        )
        workspace_pages = (
            ("Transactions & Review", review_page),
            ("Monthly Overview", overview_page),
            ("History & Trends", self.history_page),
        )
        self.workspace_tab_buttons = []
        for column, (title, page) in enumerate(workspace_pages):
            page.grid(row=0, column=0, sticky="nsew")
            button = RoundedButton(
                tab_bar,
                text=title,
                style="Tab.TButton",
                command=lambda selected_page=page: self._select_workspace_page(
                    selected_page
                ),
                height=48,
            )
            button.grid(
                row=0,
                column=column,
                sticky="ew",
                padx=(0, 8) if column < len(workspace_pages) - 1 else 0,
            )
            self.workspace_tab_buttons.append((button, page))
        self._selected_workspace_page = review_page
        review_page.tkraise()
        content = ttk.Panedwindow(review_page, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True)
        transaction_surface = RoundedCard(
            content, background=SURFACE, outline=SURFACE, width=640, height=520
        )
        review_surface = RoundedCard(
            content, background=SURFACE, outline=SURFACE, width=420, height=520
        )
        transaction_card = ttk.Frame(
            transaction_surface.content, style="Card.TFrame", padding=20
        )
        review_card = ttk.Frame(
            review_surface.content, style="Card.TFrame", padding=20
        )
        transaction_card.pack(fill=tk.BOTH, expand=True)
        review_card.pack(fill=tk.BOTH, expand=True)
        content.add(transaction_surface, weight=7)
        content.add(review_surface, weight=4)

        transaction_card.columnconfigure(0, weight=1)
        transaction_card.rowconfigure(3, weight=1)
        ttk.Label(
            transaction_card, text="Transaction Preview", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            transaction_card, textvariable=self.summary_var, style="Muted.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        filter_frame = ttk.Frame(transaction_card, style="Card.TFrame")
        filter_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_frame.columnconfigure(2, weight=1)
        filters_expanded = False
        filters_panel = ttk.Frame(filter_frame, style="Card.TFrame")
        filters_panel.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        filters_panel.grid_remove()
        for column in range(3):
            filters_panel.columnconfigure(column, weight=1)

        def update_filter_toggle():
            active = (
                self.transaction_search_var.get().strip()
                or self.transaction_category_filter_var.get() != "All Categories"
                or self.transaction_source_filter_var.get() != "All Sources"
                or self.transaction_month_filter_var.get() != "All Months"
            )
            marker = "*" if active else ""
            arrow = "▲" if filters_expanded else "▼"
            filter_toggle_button.configure(
                text=f"Search & Filters{marker} {arrow}"
            )

        def toggle_filters():
            nonlocal filters_expanded
            filters_expanded = not filters_expanded
            if filters_expanded:
                filters_panel.grid()
                search_label.grid()
                search_entry.grid()
                search_entry.focus_set()
            else:
                filters_panel.grid_remove()
                search_label.grid_remove()
                search_entry.grid_remove()
            update_filter_toggle()

        filter_toggle_button = RoundedButton(
            filter_frame,
            text="Search & Filters ▼",
            command=toggle_filters,
        )
        filter_toggle_button.grid(row=0, column=0, sticky="w")
        self._update_transaction_filter_button = update_filter_toggle

        search_label = ttk.Label(
            filter_frame, text="Search", style="Muted.TLabel"
        )
        search_label.grid(
            row=0, column=1, sticky="e", padx=(8, 8)
        )
        search_entry = ttk.Entry(
            filter_frame,
            textvariable=self.transaction_search_var,
            width=20,
        )
        search_entry.grid(row=0, column=2, sticky="ew")
        search_label.grid_remove()
        search_entry.grid_remove()
        self.transaction_category_filter = ttk.Combobox(
            filters_panel,
            textvariable=self.transaction_category_filter_var,
            values=("All Categories", "Needs Review", *self.parser.categories),
            state="readonly",
            width=15,
        )
        self.transaction_category_filter.grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        self.transaction_source_filter = ttk.Combobox(
            filters_panel,
            textvariable=self.transaction_source_filter_var,
            values=("All Sources",),
            state="readonly",
            width=11,
        )
        self.transaction_source_filter.grid(
            row=0, column=1, sticky="ew", padx=(0, 8)
        )
        self.transaction_month_filter = ttk.Combobox(
            filters_panel,
            textvariable=self.transaction_month_filter_var,
            values=("All Months",),
            state="readonly",
            width=9,
        )
        self.transaction_month_filter.grid(
            row=0, column=2, sticky="ew"
        )
        RoundedButton(
            filters_panel,
            text="Clear Filters",
            command=self._clear_transaction_filters,
        ).grid(row=0, column=3, sticky="e", padx=(8, 0))

        transaction_columns = ("date", "merchant", "amount", "category", "bank")
        self.transaction_tree = self._create_treeview(
            transaction_card,
            columns=transaction_columns,
            show="headings",
            selectmode="browse",
        )
        for column, heading, width, anchor in (
            ("date", "Date", 100, "w"),
            ("merchant", "Merchant", 260, "w"),
            ("amount", "Amount", 110, "center"),
            ("category", "Category", 150, "center"),
            ("bank", "Bank", 130, "center"),
        ):
            self.transaction_tree.heading(column, text=heading)
            self.transaction_tree.column(
                column, width=width, minwidth=65, anchor=anchor
            )
        self._transaction_columns_fitted = False
        self.transaction_tree.bind(
            "<Configure>", self._fit_initial_transaction_columns, add="+"
        )
        transaction_scroll = ttk.Scrollbar(
            transaction_card, orient=tk.VERTICAL, command=self.transaction_tree.yview
        )
        self.transaction_tree.configure(yscrollcommand=transaction_scroll.set)
        self.transaction_tree.grid(row=3, column=0, sticky="nsew")
        transaction_scroll.grid(row=3, column=1, sticky="ns")

        def filters_changed(*_):
            update_filter_toggle()
            self._refresh_transaction_table()

        for filter_combo in (
            self.transaction_category_filter,
            self.transaction_source_filter,
            self.transaction_month_filter,
        ):
            filter_combo.bind("<<ComboboxSelected>>", filters_changed)

        self.transaction_search_var.trace_add("write", filters_changed)

        review_card.columnconfigure(0, weight=1)
        review_card.rowconfigure(2, weight=1)
        ttk.Label(
            review_card, text="Review Merchants", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            review_card,
            text="Choose a category, or Ignore transfers and payments you don't want counted.",
            style="Muted.TLabel",
            wraplength=310,
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        group_columns = ("merchant", "bank", "count", "total")
        self.review_tree = self._create_treeview(
            review_card,
            columns=group_columns,
            show="headings",
            selectmode="extended",
            height=12,
        )
        self.review_tree.heading("merchant", text="Merchant")
        self.review_tree.heading("bank", text="Bank", anchor="center")
        self.review_tree.heading("count", text="#", anchor="center")
        self.review_tree.heading("total", text="Total", anchor="center")
        self.review_tree.column("merchant", width=150, minwidth=100)
        self.review_tree.column(
            "bank", width=115, minwidth=80, anchor="center"
        )
        self.review_tree.column("count", width=45, minwidth=38, anchor="center")
        self.review_tree.column("total", width=90, minwidth=70, anchor="e")
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
            values=self.parser.categories,
            state="readonly",
        )
        self.category_combo.grid(row=4, column=0, sticky="ew")
        rule_frame = ttk.Frame(review_card, style="Card.TFrame")
        rule_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        rule_frame.columnconfigure(0, weight=1)
        ttk.Label(
            rule_frame, text="Rule Match Text (Optional)", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(
            rule_frame, text="Match Type", style="CardTitle.TLabel"
        ).grid(row=0, column=1, sticky="w")
        ttk.Entry(rule_frame, textvariable=self.rule_pattern_var).grid(
            row=1, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Combobox(
            rule_frame,
            textvariable=self.rule_match_type_var,
            values=("Contains", "Exact"),
            state="readonly",
            width=9,
        ).grid(row=1, column=1, sticky="ew")
        ttk.Label(
            review_card,
            text="Leave match text blank to use the full transaction description.",
            style="Muted.TLabel",
            wraplength=310,
        ).grid(row=6, column=0, sticky="w", pady=(4, 0))
        ttk.Label(
            review_card, text="Display Title (Optional)", style="CardTitle.TLabel"
        ).grid(row=7, column=0, sticky="w", pady=(10, 5))
        ttk.Entry(review_card, textvariable=self.title_var).grid(
            row=8, column=0, sticky="ew"
        )
        RoundedButton(
            review_card,
            text="Apply To Selected Merchants",
            style="Primary.TButton",
            command=self.apply_categories,
        ).grid(row=9, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(
            review_card,
            text="Use Ctrl-click (or Command-click on Mac) to select several.",
            style="Muted.TLabel",
            wraplength=310,
        ).grid(row=10, column=0, sticky="w", pady=(8, 0))

        status_bar = ttk.Frame(shell, style="App.TFrame")
        status_bar.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(status_bar, textvariable=self.status_var, style="Subtitle.TLabel").pack(
            side=tk.LEFT
        )
        self._create_monthly_overview(overview_page)
        self._create_history_page(self.history_page)
        self._update_transaction_filter_options()
        self._on_workspace_tab_changed()

    def _select_workspace_page(self, page):
        if page is self._selected_workspace_page:
            return
        self._selected_workspace_page = page
        page.tkraise()
        self._on_workspace_tab_changed()

    def _is_history_tab_selected(self):
        return self._selected_workspace_page is self.history_page

    def _update_file_summary(self):
        count = len(self.history_files) if self._is_history_tab_selected() else len(
            self.selected_files
        )
        self.file_summary_var.set(f"Manage Files ({count})")

    def _on_workspace_tab_changed(self, _event=None):
        is_history = self._is_history_tab_selected()
        if is_history:
            self.choose_files_button.configure(
                text="Choose Excel Files", command=self.choose_history_files
            )
            self.manage_files_button.configure(command=self.manage_history_files)
            self.header_subtitle_var.set(
                "Choose or drop CombinedStatement Excel files to compare months."
            )
            self.category_targets_button.grid_remove()
            self.category_rules_button.grid_remove()
            self.export_button.grid_remove()
        else:
            self.choose_files_button.configure(
                text="Choose CSV Files", command=self.choose_files
            )
            self.manage_files_button.configure(command=self.manage_files)
            self.header_subtitle_var.set(
                "Choose or drop CSV files, review new merchants, and export your spending."
            )
            self.category_targets_button.grid()
            self.category_rules_button.grid()
            self.export_button.grid()
        for button, page in self.workspace_tab_buttons:
            button.set_selected(page is self._selected_workspace_page)
        self._update_file_summary()

    def _create_metric_card(
        self,
        parent,
        title,
        initial_value,
        accent,
        row,
        column,
        *,
        width=230,
        height=88,
        padx=0,
        pady=0,
        title_wraplength=None,
    ):
        card_surface = RoundedCard(
            parent,
            background=SURFACE,
            outline=BORDER,
            width=width,
            height=height,
            inset=8,
        )
        card_surface.grid(
            row=row,
            column=column,
            sticky="ew",
            padx=padx,
            pady=pady,
        )
        card = ttk.Frame(card_surface.content, style="Card.TFrame", padding=12)
        card.pack(fill=tk.BOTH, expand=True)
        card.columnconfigure(1, weight=1)
        accent_bar = tk.Frame(
            card,
            background=accent,
            borderwidth=0,
            highlightthickness=0,
            width=5,
        )
        accent_bar.grid(
            row=0, column=0, rowspan=2, sticky="ns", padx=(0, 12)
        )
        title_options = {"style": "Muted.TLabel"}
        if title_wraplength is not None:
            title_options["wraplength"] = title_wraplength
        ttk.Label(card, text=title, **title_options).grid(
            row=0, column=1, sticky="w"
        )
        value_var = tk.StringVar(value=initial_value)
        value_label = ttk.Label(
            card,
            textvariable=value_var,
            style="CardTitle.TLabel",
            foreground=accent,
        )
        value_label.grid(row=1, column=1, sticky="w", pady=(5, 0))
        return value_var, value_label, accent_bar

    def _set_monthly_metric_accent(self, key, color):
        self.overview_metric_labels[key].configure(foreground=color)
        self.overview_metric_accent_bars[key].configure(background=color)

    def _create_monthly_overview(self, parent):
        parent.columnconfigure(0, weight=1)
        heading = ttk.Frame(parent, style="App.TFrame")
        heading.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        heading.columnconfigure(0, weight=1)
        ttk.Label(heading, text="Monthly Overview", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(heading, text="Month", style="Subtitle.TLabel").grid(
            row=0, column=1, sticky="e", padx=(16, 8)
        )
        self.overview_month_combo = ttk.Combobox(
            heading,
            textvariable=self.overview_month_var,
            values=(),
            state="readonly",
            width=12,
        )
        self.overview_month_combo.grid(row=0, column=2, sticky="e")
        self.overview_month_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self._refresh_monthly_overview()
        )
        self.monthly_review_button = RoundedButton(
            heading,
            text="Review Duplicates",
            command=self.review_monthly_duplicates,
            state=tk.DISABLED,
        )
        self.monthly_review_button.grid(row=0, column=3, sticky="e", padx=(12, 0))

        metrics = ttk.Frame(parent, style="App.TFrame")
        metrics.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        self.overview_metric_values = {}
        self.overview_metric_labels = {}
        self.overview_metric_accent_bars = {}
        for column, (key, title, initial_value, accent) in enumerate(
            (
                ("spent", "Total Spent", "$0.00", METRIC_SPENT),
                ("budget", "Budgeted", "$0.00", METRIC_BUDGET),
                ("remaining", "Budget Remaining", "$0.00 left", STATUS_UNDER),
                ("income", "Income", "$0.00", ACCENT),
                ("net", "Net Income", "$0.00", STATUS_UNDER),
            )
        ):
            metrics.columnconfigure(
                column, weight=1, uniform="monthly-overview-metrics"
            )
            value, value_label, accent_bar = self._create_metric_card(
                metrics,
                title,
                initial_value,
                accent,
                0,
                column,
                width=200,
                padx=(0 if column == 0 else 8, 0),
            )
            self.overview_metric_values[key] = value
            self.overview_metric_labels[key] = value_label
            self.overview_metric_accent_bars[key] = accent_bar
        self.overview_spent_var = self.overview_metric_values["spent"]
        self.overview_budget_var = self.overview_metric_values["budget"]
        self.overview_remaining_var = self.overview_metric_values["remaining"]
        self.overview_income_var = self.overview_metric_values["income"]
        self.overview_net_income_var = self.overview_metric_values["net"]

        self.overview_message_var = tk.StringVar()
        ttk.Label(
            parent,
            textvariable=self.overview_message_var,
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.overview_table_surface = RoundedCard(
            parent,
            background=SURFACE,
            outline=SURFACE,
            width=800,
            height=420,
            inset=8,
        )
        self.overview_table_surface.grid(row=3, column=0, sticky="nsew")
        table_frame = ttk.Frame(
            self.overview_table_surface.content, style="Card.TFrame"
        )
        table_frame.pack(fill=tk.BOTH, expand=True)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(2, weight=1)
        self._overview_column_weights = [500, 200, 200, 400]
        self._overview_minimum_column_widths = (180, 120, 160, 180)
        self._overview_column_containers = []
        self._overview_resize_drag = None
        self._overview_header_labels = []
        self._overview_table_labels = []
        header = ttk.Frame(table_frame, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        self._overview_header_frame = header
        self._overview_column_containers = [header]
        for column, (title, anchor, weight) in enumerate(
            zip(
                ("Category", "Spent", "Monthly Target", "Budget Status"),
                ("w", "e", "e", "center"),
                self._overview_column_weights,
            )
        ):
            header.columnconfigure(
                column, weight=weight, uniform="monthly-overview"
            )
            label = ttk.Label(
                header,
                text=title,
                style="DataTable.Heading.TLabel",
                anchor=anchor,
                width=1,
                wraplength=1,
            )
            label.grid(row=0, column=column, sticky="ew")
            self._overview_header_labels.append((label, column))
            self._overview_table_labels.append((label, column))
        self._overview_column_resize_handles = []
        for boundary in range(len(self._overview_column_weights) - 1):
            handle = tk.Frame(
                header,
                background=BORDER,
                borderwidth=0,
                highlightthickness=0,
                cursor="sb_h_double_arrow",
            )
            handle.place(x=0, y=0, width=4, height=1)
            handle.bind(
                "<Enter>",
                lambda _event, grip=handle: grip.configure(background=ACCENT),
            )
            handle.bind(
                "<Leave>",
                lambda _event, grip=handle: grip.configure(background=BORDER),
            )
            handle.bind(
                "<ButtonPress-1>",
                lambda event, index=boundary: self._start_overview_column_resize(
                    index, event
                ),
            )
            handle.bind(
                "<B1-Motion>",
                lambda event, index=boundary: self._drag_overview_column_resize(
                    index, event
                ),
            )
            handle.bind("<ButtonRelease-1>", self._finish_overview_column_resize)
            self._overview_column_resize_handles.append(handle)
        header.bind("<Configure>", self._position_overview_resize_handles, add="+")
        ttk.Separator(
            table_frame,
            orient=tk.HORIZONTAL,
            style="DataTable.Horizontal.TSeparator",
        ).grid(row=1, column=0, sticky="ew")
        body = ttk.Frame(table_frame, style="Card.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.overview_canvas = tk.Canvas(
            body,
            background=SURFACE,
            highlightthickness=0,
            borderwidth=0,
        )
        overview_scroll = ttk.Scrollbar(
            body, orient=tk.VERTICAL, command=self.overview_canvas.yview
        )
        self.overview_canvas.configure(yscrollcommand=overview_scroll.set)
        self.overview_canvas.grid(row=0, column=0, sticky="nsew")
        for event_name in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.overview_canvas.bind_all(
                event_name, self._scroll_overview_rows, add="+"
            )
        overview_scroll.grid(row=0, column=1, sticky="ns")
        header_spacer = ttk.Frame(
            header,
            style="Card.TFrame",
            width=overview_scroll.winfo_reqwidth(),
        )
        header_spacer.grid(row=0, column=4, sticky="ns")
        header_spacer.grid_propagate(False)
        self._overview_header_spacer = header_spacer
        self.overview_rows_frame = ttk.Frame(
            self.overview_canvas, style="Card.TFrame"
        )
        self.overview_rows_frame.columnconfigure(0, weight=1)
        self._overview_rows_window = self.overview_canvas.create_window(
            (0, 0), window=self.overview_rows_frame, anchor="nw"
        )
        self.overview_rows_frame.bind(
            "<Configure>", self._update_overview_scroll_region
        )
        self.overview_canvas.bind("<Configure>", self._resize_overview_rows)
        parent.rowconfigure(3, weight=1)

    def _create_history_page(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        self.history_page_canvas = tk.Canvas(
            parent,
            background=BACKGROUND,
            highlightthickness=0,
            borderwidth=0,
        )
        self.history_page_scrollbar = ttk.Scrollbar(
            parent,
            orient=tk.VERTICAL,
            command=self.history_page_canvas.yview,
        )
        self.history_page_canvas.configure(
            yscrollcommand=self.history_page_scrollbar.set
        )
        self.history_page_canvas.grid(row=0, column=0, sticky="nsew")
        self.history_page_scrollbar.grid(row=0, column=1, sticky="ns")
        parent = ttk.Frame(
            self.history_page_canvas, style="App.TFrame", padding=20
        )
        self.history_page_content = parent
        self.history_page_window = self.history_page_canvas.create_window(
            (0, 0), window=parent, anchor="nw"
        )
        parent.bind("<Configure>", self._update_history_page_scroll_region)
        self.history_page_canvas.bind(
            "<Configure>", self._resize_history_page_content
        )
        for event_name in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.history_page_canvas.bind_all(
                event_name, self._scroll_history_page, add="+"
            )

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(5, weight=1)
        heading = ttk.Frame(parent, style="App.TFrame")
        heading.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        heading.columnconfigure(0, weight=1)
        ttk.Label(heading, text="History & Trends", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            heading,
            text="Compare months from previously exported CombinedStatement workbooks.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        toolbar = ttk.Frame(parent, style="App.TFrame")
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.history_review_button = RoundedButton(
            toolbar,
            text="Review Duplicates",
            command=self.review_history_duplicates,
            state=tk.DISABLED,
        )
        self.history_review_button.pack(side=tk.LEFT, padx=(8, 0))
        self.history_analyze_button = RoundedButton(
            toolbar,
            text="Analyze",
            command=self.analyze_history_data,
            state=tk.DISABLED,
        )
        self.history_analyze_button.pack(side=tk.RIGHT)
        self.history_export_button = RoundedButton(
            toolbar,
            text="Export Workbook",
            command=self.export_history_data,
            state=tk.DISABLED,
        )
        self.history_export_button.pack(side=tk.RIGHT, padx=(0, 8))

        files_surface = RoundedCard(
            parent, background=SURFACE, outline=SURFACE, width=800, height=245
        )
        files_surface.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        files_frame = ttk.Frame(files_surface.content, style="Card.TFrame", padding=12)
        files_frame.pack(fill=tk.BOTH, expand=True)
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        ttk.Label(
            files_frame, text="Source Workbooks", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Label(
            files_frame, text="Scroll to view all", style="Muted.TLabel"
        ).grid(row=0, column=1, sticky="e", pady=(0, 8))
        self.history_file_tree = self._create_treeview(
            files_frame,
            columns=("workbook", "period", "transactions"),
            show="headings",
            height=3,
            selectmode="none",
        )
        for column, title, width, anchor in (
            ("workbook", "Workbook", 440, "w"),
            ("period", "Date Range", 180, "center"),
            ("transactions", "Transactions", 110, "e"),
        ):
            self.history_file_tree.heading(column, text=title)
            self.history_file_tree.column(
                column, width=width, minwidth=80, anchor=anchor
            )
        workbook_scroll = ttk.Scrollbar(
            files_frame,
            orient=tk.VERTICAL,
            command=self.history_file_tree.yview,
        )
        self.history_file_tree.configure(yscrollcommand=workbook_scroll.set)
        self.history_file_tree.grid(row=1, column=0, sticky="nsew")
        workbook_scroll.grid(row=1, column=1, sticky="ns")
        ttk.Label(
            parent,
            textvariable=self.history_summary_var,
            style="Subtitle.TLabel",
            wraplength=1100,
        ).grid(row=3, column=0, sticky="w", pady=(0, 10))

        result_tabs = ttk.Frame(parent, style="App.TFrame")
        result_tabs.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        for column in range(4):
            result_tabs.columnconfigure(
                column, weight=1, uniform="history-tabs"
            )

        results_surface = RoundedCard(
            parent,
            background=SURFACE,
            outline=SURFACE,
            width=800,
            height=560,
        )
        self.history_results_surface = results_surface
        results_surface.grid(row=5, column=0, sticky="nsew")
        self.history_results = ttk.Frame(
            results_surface.content, style="Card.TFrame"
        )
        self.history_results.pack(fill=tk.BOTH, expand=True)
        self.history_results.columnconfigure(0, weight=1)
        self.history_results.rowconfigure(0, weight=1)
        overview_page = ttk.Frame(
            self.history_results, style="Card.TFrame", padding=10
        )
        category_page = ttk.Frame(
            self.history_results, style="Card.TFrame", padding=10
        )
        chart_page = ttk.Frame(
            self.history_results, style="Card.TFrame", padding=10
        )
        subscription_page = ttk.Frame(
            self.history_results, style="Card.TFrame", padding=10
        )
        overview_page.grid(row=0, column=0, sticky="nsew")
        category_page.grid(row=0, column=0, sticky="nsew")
        chart_page.grid(row=0, column=0, sticky="nsew")
        subscription_page.grid(row=0, column=0, sticky="nsew")
        self.history_overview_page = overview_page
        self.history_chart_page = chart_page
        self.history_category_page = category_page
        self.history_subscription_page = subscription_page
        overview_page.columnconfigure(0, weight=1)
        overview_page.rowconfigure(0, weight=1)
        self.history_overview_cards_frame = ttk.Frame(
            overview_page, style="Card.TFrame"
        )
        self.history_overview_cards_frame.grid(
            row=0, column=0, sticky="new"
        )
        for column in range(4):
            self.history_overview_cards_frame.columnconfigure(
                column, weight=1, uniform="history-overview-cards"
            )
        ttk.Label(
            overview_page,
            text=(
                "Budget totals use current monthly targets across the months "
                "analyzed. Category averages start at the first month of spending."
            ),
            style="Muted.TLabel",
            wraplength=1100,
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
        category_page.columnconfigure(0, weight=1)
        category_page.rowconfigure(0, weight=1)
        chart_page.columnconfigure(0, weight=1)
        chart_page.rowconfigure(0, weight=1)
        self.history_category_chart = tk.Canvas(
            chart_page,
            background=SURFACE,
            height=1,
            highlightthickness=0,
            borderwidth=0,
        )
        self.history_category_chart.grid(row=0, column=0, sticky="nsew")
        self.history_category_chart.bind(
            "<Configure>",
            lambda _event: self._draw_history_category_chart(),
            add="+",
        )
        self.history_result_tab_buttons = []
        for column, (title, page) in enumerate(
            (
                ("Overview", overview_page),
                ("Trend Graph", chart_page),
                ("Category Trends", category_page),
                ("Likely Subscriptions", subscription_page),
            )
        ):
            button = RoundedButton(
                result_tabs,
                text=title,
                style="Tab.TButton",
                command=lambda selected_page=page: self._select_history_results_page(
                    selected_page
                ),
                height=42,
            )
            button.grid(
                row=0,
                column=column,
                sticky="ew",
                padx=(0, 8) if column < 3 else 0,
            )
            self.history_result_tab_buttons.append((button, page))
        self._selected_history_results_page = overview_page
        overview_page.tkraise()
        self.history_result_tab_buttons[0][0].set_selected(True)
        self.history_category_tree = self._create_history_tree(
            category_page,
            (
                ("category", "Category", 240, "w"),
                ("total", "Total Spent", 160, "e"),
                ("average", "Average / Month Since First Spend", 230, "e"),
                ("change", "Change / Month", 160, "e"),
                ("percent", "% Change vs. Avg.", 165, "e"),
                ("direction", "Trend", 150, "center"),
            ),
        )
        self.history_category_tree.configure(style="CategoryTrend.Treeview")
        self.history_category_tree.tag_configure(
            "increasing", foreground=STATUS_OVER
        )
        self.history_category_tree.tag_configure(
            "decreasing", foreground=STATUS_UNDER
        )
        self.history_category_tree.tag_configure(
            "stable", foreground=TEXT_SECONDARY
        )
        self.history_category_colors = {}
        self.history_subscription_tree = self._create_history_tree(
            subscription_page,
            (
                ("merchant", "Merchant", 290, "w"),
                ("amount", "Amount", 120, "e"),
                ("months", "Months Seen", 110, "e"),
                ("first", "First Seen", 130, "center"),
                ("last", "Last Seen", 130, "center"),
                ("count", "Transactions", 120, "e"),
            ),
        )
        self._refresh_history_overview_cards()
        self._update_history_results_size()

    def _resize_history_page_content(self, event):
        self.history_page_canvas.itemconfigure(
            self.history_page_window, width=event.width
        )
        self._update_history_page_scroll_region()

    def _update_history_page_scroll_region(self, _event=None):
        self.history_page_canvas.configure(
            scrollregion=self.history_page_canvas.bbox("all")
        )

    def _scroll_history_page(self, event):
        widget = event.widget
        while widget is not None and widget is not self.history_page_canvas:
            widget = getattr(widget, "master", None)
        if (
            widget is None
            or self._selected_workspace_page is not self.history_page
        ):
            return
        delta = getattr(event, "delta", 0)
        if delta:
            units = -int(delta / 120) if abs(delta) >= 120 else -int(delta)
            if units == 0:
                units = -1 if delta > 0 else 1
        else:
            units = -1 if getattr(event, "num", None) == 4 else 1
        self.history_page_canvas.yview_scroll(units, "units")
        return "break"

    def _update_history_results_size(self):
        analysis = self.history_analysis
        category_count = len(analysis.category_history) if analysis else 0
        subscription_count = len(analysis.subscriptions) if analysis else 0
        self.history_category_tree.configure(height=max(3, category_count))
        self.history_subscription_tree.configure(height=max(3, subscription_count))
        row_count = max(8, category_count, subscription_count)
        overview_rows = (7 + category_count + 3) // 4
        self.history_results_surface.configure(
            height=max(560, row_count * 48 + 120, overview_rows * 128 + 60)
        )

    def _refresh_history_overview_cards(self):
        for widget in self.history_overview_cards_frame.winfo_children():
            widget.destroy()
        self.history_overview_metric_values = {}

        analysis = self.history_analysis
        month_count = len(analysis.months) if analysis else 0
        total_spent = analysis.total_spending if analysis else 0.0
        total_income = analysis.total_income if analysis else 0.0
        monthly_budget = sum(
            amount
            for category, amount in self.target_handler.targets.items()
            if category not in (Category.INCOME.value, Category.IGNORE.value)
        )
        total_budgeted = monthly_budget * month_count
        net_total = total_income - total_spent
        unbudgeted_income = total_income - total_budgeted
        average_spent = total_spent / month_count if month_count else 0.0
        average_income = total_income / month_count if month_count else 0.0

        cards = [
            ("Total Spent", total_spent, METRIC_SPENT),
            ("Total Income", total_income, ACCENT),
            (
                "Net Total",
                net_total,
                STATUS_UNDER if net_total >= 0 else STATUS_OVER,
            ),
            ("Total Budgeted", total_budgeted, METRIC_BUDGET),
            (
                "Unbudgeted Income",
                unbudgeted_income,
                STATUS_UNDER if unbudgeted_income >= 0 else STATUS_OVER,
            ),
            ("Average Monthly Spent", average_spent, METRIC_SPENT),
            ("Average Monthly Income", average_income, ACCENT),
        ]
        if analysis:
            cards.extend(
                (
                    f"{category.category} Avg / Month",
                    category.average_monthly_spend,
                    self.history_category_colors.get(category.category, ACCENT),
                )
                for category in analysis.category_history
            )

        for index, (title, value, accent) in enumerate(cards):
            value_var, _, _ = self._create_metric_card(
                self.history_overview_cards_frame,
                title,
                self._format_amount(value),
                accent,
                index // 4,
                index % 4,
                width=250,
                height=112,
                padx=8,
                pady=8,
                title_wraplength=220,
            )
            self.history_overview_metric_values[title] = value_var

    def _select_history_results_page(self, page):
        if page is self._selected_history_results_page:
            return
        self._selected_history_results_page = page
        page.tkraise()
        if page is self.history_overview_page:
            self._refresh_history_overview_cards()
            self._update_history_results_size()
        if page is self.history_chart_page:
            self.after_idle(self._draw_history_category_chart)
        for button, tab_page in self.history_result_tab_buttons:
            button.set_selected(tab_page is page)

    def _create_history_tree(self, parent, columns):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        tree = self._create_treeview(
            parent,
            columns=tuple(column[0] for column in columns),
            show="headings",
        )
        for column, heading, width, anchor in columns:
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=80, anchor=anchor)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        return tree

    def _create_treeview(self, *args, **kwargs):
        tree = ttk.Treeview(*args, **kwargs)
        columns = tree.tk.splitlist(tree.cget("columns"))
        handles = []
        for boundary in range(len(columns) - 1):
            handle = tk.Frame(
                tree,
                background=BORDER,
                borderwidth=0,
                highlightthickness=0,
                cursor="sb_h_double_arrow",
            )
            handle.bind(
                "<Enter>",
                lambda _event, grip=handle: grip.configure(background=ACCENT),
            )
            handle.bind(
                "<Leave>",
                lambda _event, grip=handle: grip.configure(background=BORDER),
            )
            handle.bind(
                "<ButtonPress-1>",
                lambda event, widget=tree, index=boundary: self._start_treeview_column_resize(
                    widget, index, event
                ),
            )
            handle.bind(
                "<B1-Motion>",
                lambda event, widget=tree, index=boundary: self._drag_treeview_column_resize(
                    widget, index, event
                ),
            )
            handle.bind(
                "<ButtonRelease-1>",
                lambda event, widget=tree: self._finish_treeview_column_resize(
                    widget, event
                ),
            )
            handles.append(handle)
        tree._column_resize_handles = handles
        tree._column_resize_drag = None
        tree.bind(
            "<Configure>",
            lambda _event, widget=tree: self._position_treeview_resize_handles(
                widget
            ),
            add="+",
        )
        tree.bind(
            "<B1-Motion>",
            lambda _event, widget=tree: self._schedule_treeview_handle_position(
                widget
            ),
            add="+",
        )
        tree.bind(
            "<ButtonRelease-1>",
            lambda _event, widget=tree: self._schedule_treeview_handle_position(
                widget
            ),
            add="+",
        )
        tree.after_idle(lambda widget=tree: self._position_treeview_resize_handles(widget))
        return tree

    def _fit_initial_transaction_columns(self, _event=None):
        if self._transaction_columns_fitted:
            return
        tree = self.transaction_tree
        available_width = tree.winfo_width()
        if available_width <= 1:
            return

        self._transaction_columns_fitted = True
        columns = tree.tk.splitlist(tree.cget("columns"))
        widths = {
            column: int(tree.column(column, "width")) for column in columns
        }
        deficit = max(0, sum(widths.values()) - available_width)
        for column in ("merchant", "category", "amount", "date", "bank"):
            reduction = min(
                deficit,
                widths[column] - int(tree.column(column, "minwidth")),
            )
            widths[column] -= reduction
            deficit -= reduction
            if deficit == 0:
                break
        for column, width in widths.items():
            tree.column(column, width=width)

    def _position_treeview_resize_handles(self, tree):
        columns = tree.tk.splitlist(tree.cget("columns"))
        x_position = 0
        tree_width = tree.winfo_width()
        header_top = 0
        while (
            header_top < tree.winfo_height()
            and tree.identify_region(8, header_top) != "heading"
        ):
            header_top += 1
        header_bottom = header_top
        while (
            header_bottom < tree.winfo_height()
            and tree.identify_region(8, header_bottom) == "heading"
        ):
            header_bottom += 1

        for column, handle in zip(columns[:-1], tree._column_resize_handles):
            x_position += int(tree.column(column, "width"))
            if 0 < x_position < tree_width and header_bottom > header_top:
                handle.place(
                    x=x_position - 2,
                    y=header_top,
                    width=4,
                    height=header_bottom - header_top,
                )
            else:
                handle.place_forget()

    def _schedule_treeview_handle_position(self, tree):
        tree.after_idle(lambda widget=tree: self._position_treeview_resize_handles(widget))

    def _start_treeview_column_resize(self, tree, boundary, event):
        tree._column_resize_drag = boundary
        event.widget.configure(background=ACCENT)
        self._forward_treeview_resize_event(tree, "<ButtonPress-1>", event)

    def _drag_treeview_column_resize(self, tree, boundary, event):
        if tree._column_resize_drag != boundary:
            return
        self._forward_treeview_resize_event(tree, "<B1-Motion>", event)

    def _finish_treeview_column_resize(self, tree, event):
        tree._column_resize_drag = None
        event.widget.configure(background=BORDER)
        self._forward_treeview_resize_event(tree, "<ButtonRelease-1>", event)

    @staticmethod
    def _forward_treeview_resize_event(tree, sequence, event):
        tree.event_generate(
            sequence,
            x=event.x_root - tree.winfo_rootx(),
            y=event.y_root - tree.winfo_rooty(),
            rootx=event.x_root,
            rooty=event.y_root,
            state=getattr(event, "state", 0),
            when="now",
        )

    def choose_history_files(self, parent=None):
        paths = filedialog.askopenfilenames(
            title="Add CombinedStatement Workbooks",
            filetypes=(
                ("CombinedStatement workbooks", "*.xlsx"),
                ("Excel workbooks", "*.xlsm"),
            ),
            parent=self if parent is None else parent,
        )
        if paths:
            self._add_history_files(paths)

    def _add_history_files(self, paths):
        known_paths = {self._file_identity(path) for path in self.history_files}
        errors = []
        added = []
        for path in paths:
            identity = self._file_identity(path)
            if identity in known_paths:
                continue
            known_paths.add(identity)
            try:
                statement = load_combined_statement(path)
            except (
                OSError,
                ValueError,
                KeyError,
                BadZipFile,
                InvalidFileException,
            ) as error:
                errors.append(f"{os.path.basename(path)}: {error}")
                continue
            self.history_files.append(path)
            self.history_statements.append(statement)
            added.append(path)
        if added:
            self.history_analysis = None
            self.history_duplicates_reviewed = False
            self._refresh_history()
            self.status_var.set(f"Added {len(added)} historical workbook(s).")
        if errors:
            messagebox.showwarning(
                "Some Workbooks Could Not Be Added",
                "\n\n".join(errors),
                parent=self,
            )
        if not added and not errors and paths:
            messagebox.showinfo(
                "Workbooks Already Loaded",
                "Those workbooks are already included in the history analysis.",
                parent=self,
            )

    def _remove_history_files_at(self, remove_indexes):
        if not remove_indexes:
            return
        self.history_files = [
            path
            for index, path in enumerate(self.history_files)
            if index not in remove_indexes
        ]
        self.history_statements = [
            statement
            for index, statement in enumerate(self.history_statements)
            if index not in remove_indexes
        ]
        self.history_duplicate_exclusions.clear()
        self.history_analysis = None
        self.history_duplicates_reviewed = False
        self._refresh_history()

    def manage_history_files(self):
        dialog = tk.Toplevel(self)
        dialog.configure(background=BACKGROUND)
        dialog.title("Manage History Workbooks")
        dialog.transient(self)
        dialog.geometry("900x600")
        dialog.minsize(680, 480)
        frame = ttk.Frame(dialog, style="App.TFrame", padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text="Manage History Workbooks", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text=(
                "Add CombinedStatement workbooks or remove them from this history "
                "analysis. Source files on disk are not changed."
            ),
            style="Subtitle.TLabel",
            wraplength=800,
        ).grid(row=1, column=0, sticky="w", pady=(5, 14))
        file_tree = self._create_treeview(
            frame,
            columns=("path",),
            show="headings",
            selectmode="extended",
            height=6,
        )
        file_tree.heading("path", text="Selected Excel Workbooks")
        file_tree.column("path", width=820, minwidth=300)
        file_scroll = ttk.Scrollbar(
            frame, orient=tk.VERTICAL, command=file_tree.yview
        )
        file_tree.configure(yscrollcommand=file_scroll.set)
        file_tree.grid(row=2, column=0, sticky="nsew")
        file_scroll.grid(row=2, column=1, sticky="ns")

        def refresh_files():
            for item in file_tree.get_children():
                file_tree.delete(item)
            for index, path in enumerate(self.history_files):
                file_tree.insert("", tk.END, iid=f"file-{index}", values=(path,))

        def add_files():
            self.choose_history_files(parent=dialog)
            refresh_files()

        def remove_files():
            selection = file_tree.selection()
            if not selection:
                messagebox.showinfo(
                    "Select Workbooks",
                    "Select one or more workbooks to remove from this analysis.",
                    parent=dialog,
                )
                return
            if not messagebox.askyesno(
                "Remove Selected Workbooks?",
                "Remove the selected workbooks and their transactions from this "
                "history analysis? The original files will not be deleted.",
                parent=dialog,
            ):
                return
            remove_indexes = {
                int(item.removeprefix("file-")) for item in selection
            }
            self._remove_history_files_at(remove_indexes)
            refresh_files()

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        RoundedButton(
            buttons,
            text="Add Excel Files",
            style="Primary.TButton",
            command=add_files,
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons,
            text="Remove Selected",
            command=remove_files,
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(buttons, text="Done", command=dialog.destroy).pack(
            side=tk.RIGHT
        )
        refresh_files()
        self._center_dialog(dialog)

    def _refresh_history(self):
        self.history_duplicate_groups = find_duplicate_groups(
            self.history_statements, self.history_files
        )
        policy = self.parser.duplicate_policy
        self.history_effective_exclusions = set()
        if policy == "warn_review":
            self.history_effective_exclusions.update(
                self.history_duplicate_exclusions
            )
        elif policy == "auto_merge":
            self.history_effective_exclusions.update(
                self.history_duplicate_exclusions
            )
            self.history_effective_exclusions.update(
                duplicate_occurrences_to_skip(self.history_duplicate_groups)
            )
        self.history_analysis_statements = filter_duplicate_occurrences(
            self.history_statements, self.history_effective_exclusions
        )
        for item in self.history_file_tree.get_children():
            self.history_file_tree.delete(item)
        for index, (path, statement) in enumerate(
            zip(self.history_files, self.history_statements)
        ):
            dates = [transaction.date for transaction in statement.transactions]
            period = (
                f"{min(dates):%Y-%m-%d} to {max(dates):%Y-%m-%d}"
                if dates
                else "No dates"
            )
            self.history_file_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(os.path.basename(path), period, len(statement.transactions)),
            )
        self._update_history_controls()
        if self.history_analysis is None:
            self._clear_history_results()
            if not self.history_files:
                self.history_summary_var.set(
                    "Add CombinedStatement workbooks to analyze multiple months."
                )
            elif self.history_duplicate_groups:
                self.history_summary_var.set(
                    f"{len(self.history_files)} workbook(s) loaded; "
                    f"{len(self.history_duplicate_groups)} potential duplicate group(s). "
                    "Analyze to view category trends and recurring transactions."
                )
            else:
                self.history_summary_var.set(
                    f"{len(self.history_files)} workbook(s) loaded. "
                    "Analyze to view category trends and recurring transactions."
                )

    def _update_history_controls(self):
        has_files = bool(self.history_files)
        self.history_analyze_button.configure(
            state=tk.NORMAL if has_files else tk.DISABLED
        )
        self.history_export_button.configure(
            state=tk.NORMAL if self.history_analysis is not None else tk.DISABLED
        )
        duplicate_count = len(self.history_duplicate_groups)
        policy = self.parser.duplicate_policy
        self.history_review_button.configure(
            text=f"Review Duplicates ({duplicate_count})",
            state=tk.NORMAL
            if duplicate_count and policy == "warn_review"
            else tk.DISABLED,
        )
        self._update_file_summary()

    def _clear_history_results(self):
        for tree in (self.history_category_tree, self.history_subscription_tree):
            for item in tree.get_children():
                tree.delete(item)
        self._refresh_history_overview_cards()
        self._update_history_results_size()
        self._draw_history_category_chart()
        self.history_export_button.configure(state=tk.DISABLED)

    def analyze_history_data(self):
        if not self.history_statements:
            return
        if (
            self.history_duplicate_groups
            and self.parser.duplicate_policy == "warn_review"
            and not self.history_duplicates_reviewed
        ):
            review_now = messagebox.askyesnocancel(
                "Potential Duplicates",
                "Potential duplicate transactions were found. Review them before "
                "analyzing? Choose No to include all matching transactions.",
                parent=self,
            )
            if review_now is None:
                return
            if review_now:
                self.review_history_duplicates()
            else:
                self.history_duplicates_reviewed = True
        self._refresh_history()
        self.history_analysis = analyze_history(self.history_analysis_statements)
        if not self.history_analysis.months:
            messagebox.showinfo(
                "No Historical Transactions",
                "The selected workbooks do not contain transactions to analyze.",
                parent=self,
            )
            self.history_analysis = None
            self._clear_history_results()
            self._update_history_controls()
            return
        for item in self.history_category_tree.get_children():
            self.history_category_tree.delete(item)
        for item in self.history_subscription_tree.get_children():
            self.history_subscription_tree.delete(item)
        for index, category in enumerate(self.history_analysis.category_history):
            change = category.monthly_change
            if change is None:
                direction = "Need 2+ months"
                trend_tag = "stable"
            elif change > 0.01:
                direction = "Increasing"
                trend_tag = "increasing"
            elif change < -0.01:
                direction = "Decreasing"
                trend_tag = "decreasing"
            else:
                direction = "Stable"
                trend_tag = "stable"
            average = category.average_monthly_spend
            percentage_change = (
                f"{change / average:+.1%}"
                if change is not None and average > 0
                else "—"
            )
            self.history_category_tree.insert(
                "",
                tk.END,
                iid=f"history-category-{index}",
                values=(
                    category.category,
                    self._format_amount(sum(category.monthly_spending)),
                    self._format_amount(average),
                    "—" if change is None else self._format_amount(change),
                    percentage_change,
                    direction,
                ),
                tags=(trend_tag,),
            )
        for index, subscription in enumerate(self.history_analysis.subscriptions):
            self.history_subscription_tree.insert(
                "",
                tk.END,
                iid=f"history-subscription-{index}",
                values=(
                    subscription.merchant,
                    self._format_amount(subscription.amount),
                    len(subscription.months),
                    subscription.first_date.strftime("%Y-%m-%d"),
                    subscription.last_date.strftime("%Y-%m-%d"),
                    subscription.transaction_count,
                ),
            )
        self._update_history_results_size()
        months = self.history_analysis.months
        self._draw_history_category_chart()
        self._refresh_history_overview_cards()
        self.history_summary_var.set(
            f"Analyzed {len(months)} month(s), "
            f"{sum(len(statement.transactions) for statement in self.history_analysis_statements)} "
            f"transaction(s), and found {len(self.history_analysis.subscriptions)} "
            "likely subscription(s)."
        )
        self._update_history_controls()

    def _show_history_chart_tooltip(self, event, label, average, color):
        canvas = self.history_category_chart
        canvas.delete("history-chart-tooltip")
        text_item = canvas.create_text(
            event.x + 14,
            event.y + 14,
            text=f"{label} average: {self._format_amount(average)} / month",
            anchor="nw",
            fill=TEXT_PRIMARY,
            font=("Arial", 11, "bold"),
            width=max(100, min(420, canvas.winfo_width() - 28)),
            tags=("history-chart-tooltip", "history-chart-tooltip-text"),
        )
        text_bbox = canvas.bbox(text_item)
        if text_bbox is None:
            return
        rectangle = canvas.create_rectangle(
            text_bbox[0] - 8,
            text_bbox[1] - 5,
            text_bbox[2] + 8,
            text_bbox[3] + 5,
            fill=SURFACE_RAISED,
            outline=color,
            width=1,
            tags=("history-chart-tooltip", "history-chart-tooltip-background"),
        )
        canvas.tag_lower(rectangle, text_item)
        self._position_history_chart_tooltip(event)

    def _position_history_chart_tooltip(self, event):
        canvas = self.history_category_chart
        text_items = canvas.find_withtag("history-chart-tooltip-text")
        background_items = canvas.find_withtag("history-chart-tooltip-background")
        if not text_items or not background_items:
            return
        text_item = text_items[0]
        text_bbox = canvas.bbox(text_item)
        if text_bbox is None:
            return
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        x = max(
            6,
            min(event.x + 14, canvas.winfo_width() - text_width - 22),
        )
        y = max(
            6,
            min(event.y + 14, canvas.winfo_height() - text_height - 18),
        )
        canvas.coords(text_item, x + 8, y + 5)
        text_bbox = canvas.bbox(text_item)
        if text_bbox is None:
            return
        canvas.coords(
            background_items[0],
            text_bbox[0] - 8,
            text_bbox[1] - 5,
            text_bbox[2] + 8,
            text_bbox[3] + 5,
        )
        canvas.tag_lower(background_items[0], text_item)

    def _hide_history_chart_tooltip(self, _event=None):
        self.history_category_chart.delete("history-chart-tooltip")

    def _draw_history_category_chart(self):
        canvas = self.history_category_chart
        canvas.delete("all")
        analysis = self.history_analysis
        categories = analysis.category_history if analysis else ()
        months = analysis.months if analysis else ()

        category_count = len(categories)
        self.history_category_colors = {}
        for index, category in enumerate(categories):
            self.history_category_visibility.setdefault(category.category, True)
            hue = (index / max(1, category_count) + 0.04) % 1
            red, green, blue = colorsys.hsv_to_rgb(hue, 0.68, 0.95)
            self.history_category_colors[category.category] = (
                f"#{round(red * 255):02x}{round(green * 255):02x}"
                f"{round(blue * 255):02x}"
            )

        width = canvas.winfo_width()
        height = canvas.winfo_height()
        if width <= 1 or height <= 1:
            return
        if not categories or not months:
            message = (
                "Analyze workbooks to see category spending trends."
                if analysis is None
                else "No categorized spending available to chart."
            )
            canvas.create_text(
                width // 2,
                height // 2,
                text=message,
                fill=TEXT_SECONDARY,
                font=("Arial", 13),
            )
            return

        legend_width = min(520, max(320, int(width * 0.34)))
        legend_left = width - legend_width
        left = 68
        right = legend_left - 22
        top = 50
        bottom = height - 46
        plot_width = max(1, right - left)
        plot_height = max(1, bottom - top)
        canvas.create_text(
            left,
            12,
            text="Monthly spend (solid) and linear trend (dashed)",
            anchor="nw",
            fill=TEXT_SECONDARY,
            font=("Arial", 12),
        )
        canvas.create_text(
            legend_left + 8,
            10,
            text="Categories — click to toggle",
            anchor="nw",
            fill=TEXT_SECONDARY,
            font=("Arial", 16, "bold"),
        )

        visible_categories = tuple(
            category
            for category in categories
            if self.history_category_visibility.get(category.category, True)
        )
        total_monthly_spending = tuple(
            sum(category.monthly_spending[index] for category in visible_categories)
            for index in range(len(months))
        )
        total_first_index = next(
            (
                index
                for index, value in enumerate(total_monthly_spending)
                if value > 0
            ),
            None,
        )
        total_active_values = (
            total_monthly_spending[total_first_index:]
            if total_first_index is not None
            else ()
        )
        total_average = (
            sum(total_active_values) / len(total_active_values)
            if total_active_values
            else 0.0
        )
        total_midpoint = (len(total_active_values) - 1) / 2
        total_change = None
        if len(total_active_values) >= 2:
            denominator = sum(
                (index - total_midpoint) ** 2
                for index in range(len(total_active_values))
            )
            total_change = sum(
                (index - total_midpoint) * (value - total_average)
                for index, value in enumerate(total_active_values)
            ) / denominator
        trend_values = []
        if total_change is not None and total_first_index is not None:
            trend_values.extend(
                total_average
                + total_change * (index - total_first_index - total_midpoint)
                for index in range(total_first_index, len(months))
            )
        for category in visible_categories:
            if category.monthly_change is None:
                continue
            first_index = category.first_spending_month_index
            midpoint = (len(months) - first_index - 1) / 2
            trend_values.extend(
                category.average_monthly_spend
                + category.monthly_change * (index - first_index - midpoint)
                for index in range(first_index, len(months))
            )
        maximum = max(
            0.0,
            max(total_monthly_spending, default=0.0),
            max(
                (
                    max(category.monthly_spending, default=0.0)
                    for category in visible_categories
                ),
                default=0.0,
            ),
            max(trend_values, default=0.0),
        )
        maximum = maximum or 1.0
        for tick in range(5):
            value = maximum * tick / 4
            y = bottom - plot_height * tick / 4
            canvas.create_line(
                left,
                y,
                right,
                y,
                fill=BORDER,
                dash=(2, 4),
            )
            canvas.create_text(
                left - 8,
                y,
                text=f"${value:,.0f}",
                anchor="e",
                fill=TEXT_SECONDARY,
                font=("Arial", 10),
                tags=("history-chart-y-tick",),
            )

        def x_for(index):
            if len(months) == 1:
                return left + plot_width / 2
            return left + plot_width * index / (len(months) - 1)

        def y_for(value):
            return bottom - plot_height * max(0.0, value) / maximum

        tick_step = max(1, (len(months) + 4) // 5)
        month_ticks = list(range(0, len(months), tick_step))
        if month_ticks[-1] != len(months) - 1:
            month_ticks.append(len(months) - 1)
        for index in month_ticks:
            canvas.create_text(
                x_for(index),
                bottom + 16,
                text=months[index].strftime("%b %Y"),
                anchor="n",
                fill=TEXT_SECONDARY,
                font=("Arial", 10),
            )

        for category_index, category in enumerate(categories):
            color = self.history_category_colors[category.category]
            enabled = self.history_category_visibility.get(category.category, True)
            if enabled:
                first_index = category.first_spending_month_index
                actual_points = [
                    coordinate
                    for index, value in enumerate(category.monthly_spending)
                    if index >= first_index
                    for coordinate in (x_for(index), y_for(value))
                ]
                if len(actual_points) >= 4:
                    canvas.create_line(
                        *actual_points,
                        fill=color,
                        width=3,
                        tags=("category-series", f"category-{category_index}"),
                    )
                for index in range(first_index, len(category.monthly_spending)):
                    value = category.monthly_spending[index]
                    x = x_for(index)
                    y = y_for(value)
                    canvas.create_oval(
                        x - 3,
                        y - 3,
                        x + 3,
                        y + 3,
                        fill=color,
                        outline=color,
                    )

                if category.monthly_change is not None and len(months) >= 2:
                    midpoint = (len(months) - first_index - 1) / 2
                    trend_points = [
                        coordinate
                        for index in range(first_index, len(months))
                        for coordinate in (
                            x_for(index),
                            y_for(
                                category.average_monthly_spend
                                + category.monthly_change
                                * (index - first_index - midpoint)
                            ),
                        )
                    ]
                    trend_tag = f"history-category-trend-{category_index}"
                    canvas.create_line(
                        *trend_points,
                        fill=color,
                        width=3,
                        dash=(7, 4),
                        tags=(
                            "category-trend",
                            f"category-{category_index}",
                            trend_tag,
                        ),
                    )
                    canvas.tag_bind(
                        trend_tag,
                        "<Enter>",
                        lambda event, name=category.category, average=category.average_monthly_spend, trend_color=color: self._show_history_chart_tooltip(
                            event, name, average, trend_color
                        ),
                    )
                    canvas.tag_bind(
                        trend_tag,
                        "<Motion>",
                        self._position_history_chart_tooltip,
                    )
                    canvas.tag_bind(
                        trend_tag,
                        "<Leave>",
                        self._hide_history_chart_tooltip,
                    )

            legend_tag = f"history-category-legend-{category_index}"
            legend_x = legend_left + 10
            legend_y = top + (category_index + 1) * 48 + 14
            canvas.create_rectangle(
                legend_x,
                legend_y - 12,
                legend_x + 24,
                legend_y + 12,
                fill=color if enabled else SURFACE,
                outline=color,
                width=2,
                tags=("category-legend-toggle", legend_tag),
            )
            if enabled:
                canvas.create_line(
                    legend_x + 5,
                    legend_y,
                    legend_x + 10,
                    legend_y + 5,
                    legend_x + 19,
                    legend_y - 6,
                    fill=BACKGROUND,
                    width=3,
                    tags=("category-legend-toggle", legend_tag),
                )
            label_x = legend_x + 36
            legend_line_x = width - 62
            max_label_length = max(5, int((legend_line_x - label_x - 12) / 10))
            legend_label = category.category
            if len(legend_label) > max_label_length:
                legend_label = legend_label[: max_label_length - 3] + "..."
            canvas.create_text(
                label_x,
                legend_y,
                text=legend_label,
                anchor="w",
                fill=TEXT_PRIMARY if enabled else TEXT_SECONDARY,
                font=("Arial", 18, "bold" if enabled else "normal"),
                tags=("category-legend-label", legend_tag),
            )
            canvas.create_line(
                legend_line_x,
                legend_y,
                legend_line_x + 14,
                legend_y,
                fill=color,
                width=4,
                tags=("category-legend-line", legend_tag),
            )
            canvas.create_line(
                legend_line_x + 14,
                legend_y,
                legend_line_x + 28,
                legend_y,
                fill=color,
                width=4,
                dash=(5, 3),
                tags=("category-legend-line", legend_tag),
            )
            canvas.tag_bind(
                legend_tag,
                "<Button-1>",
                lambda _event, name=category.category: self._toggle_history_category(
                    name
                ),
            )
            canvas.tag_bind(
                legend_tag,
                "<Enter>",
                lambda _event: canvas.configure(cursor="hand2"),
            )
            canvas.tag_bind(
                legend_tag,
                "<Leave>",
                lambda _event: canvas.configure(cursor=""),
            )

        total_actual_points = (
            [
                coordinate
                for index in range(total_first_index, len(total_monthly_spending))
                for coordinate in (
                    x_for(index),
                    y_for(total_monthly_spending[index]),
                )
            ]
            if total_first_index is not None
            else []
        )
        if len(total_actual_points) >= 4:
            canvas.create_line(
                *total_actual_points,
                fill=METRIC_SPENT,
                width=4,
                tags=("total-spend-series",),
            )
        if total_first_index is not None:
            for index in range(total_first_index, len(total_monthly_spending)):
                x = x_for(index)
                y = y_for(total_monthly_spending[index])
                canvas.create_oval(
                    x - 4,
                    y - 4,
                    x + 4,
                    y + 4,
                    fill=METRIC_SPENT,
                    outline=METRIC_SPENT,
                )

        if (
            total_change is not None
            and total_first_index is not None
            and len(months) >= 2
        ):
            total_trend_points = [
                coordinate
                for index in range(total_first_index, len(months))
                for coordinate in (
                    x_for(index),
                    y_for(
                        total_average
                        + total_change
                        * (index - total_first_index - total_midpoint)
                    ),
                )
            ]
            total_trend_tag = "history-total-spend-trend"
            canvas.create_line(
                *total_trend_points,
                fill=METRIC_SPENT,
                width=4,
                dash=(10, 5),
                tags=("total-spend-trend", total_trend_tag),
            )
            canvas.tag_bind(
                total_trend_tag,
                "<Enter>",
                lambda event: self._show_history_chart_tooltip(
                    event, "Selected Total", total_average, METRIC_SPENT
                ),
            )
            canvas.tag_bind(
                total_trend_tag,
                "<Motion>",
                self._position_history_chart_tooltip,
            )
            canvas.tag_bind(
                total_trend_tag,
                "<Leave>",
                self._hide_history_chart_tooltip,
            )

        total_legend_y = top + 14
        canvas.create_text(
            legend_left + 10,
            total_legend_y,
            text="Selected Total",
            anchor="w",
            fill=METRIC_SPENT,
            font=("Arial", 18, "bold"),
            tags=("total-spend-legend-label",),
        )
        total_legend_line_x = width - 62
        canvas.create_line(
            total_legend_line_x,
            total_legend_y,
            total_legend_line_x + 14,
            total_legend_y,
            fill=METRIC_SPENT,
            width=4,
            tags=("total-spend-legend-line",),
        )
        canvas.create_line(
            total_legend_line_x + 14,
            total_legend_y,
            total_legend_line_x + 28,
            total_legend_y,
            fill=METRIC_SPENT,
            width=4,
            dash=(5, 3),
            tags=("total-spend-legend-line",),
        )

    def _toggle_history_category(self, category):
        self.history_category_visibility[category] = not (
            self.history_category_visibility.get(category, True)
        )
        self._draw_history_category_chart()

    def review_history_duplicates(self):
        self._show_duplicate_review(
            self.history_statements,
            self.history_files,
            self.history_duplicate_exclusions,
            self._refresh_history,
            "History Duplicate Review",
        )
        self.history_duplicates_reviewed = True
        self.history_analysis = None
        self._refresh_history()

    def _show_duplicate_review(
        self, statements, source_files, exclusions, on_change, title
    ):
        groups = find_duplicate_groups(statements, source_files)
        if not groups:
            messagebox.showinfo(
                "No Potential Duplicates",
                "No matching transactions were found across different files.",
                parent=self,
            )
            return

        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(background=BACKGROUND)
        dialog.geometry("1050x580")
        dialog.minsize(800, 420)
        dialog.transient(self)
        frame = ttk.Frame(dialog, style="App.TFrame", padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text=title, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text=(
                "Matches share the same date, amount, and place across different "
                "source files. Select any occurrence to exclude or restore it."
            ),
            style="Subtitle.TLabel",
            wraplength=950,
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        tree_frame = ttk.Frame(frame, style="App.TFrame")
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        tree = self._create_treeview(
            tree_frame,
            columns=("date", "amount", "place", "source", "status"),
            show="headings",
            selectmode="extended",
        )
        for column, heading, width, anchor in (
            ("date", "Date", 110, "center"),
            ("amount", "Amount", 110, "e"),
            ("place", "Place", 280, "w"),
            ("source", "Source Workbook", 380, "w"),
            ("status", "Status", 110, "center"),
        ):
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=80, anchor=anchor)
        scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=tree.yview
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        occurrences = {}
        for group in groups:
            for occurrence in group.occurrences:
                key = (
                    occurrence.statement_index,
                    occurrence.transaction_index,
                )
                occurrences[key] = occurrence

        def refresh():
            for item in tree.get_children():
                tree.delete(item)
            for (statement_index, transaction_index), occurrence in sorted(
                occurrences.items(),
                key=lambda item: (
                    item[1].transaction.date,
                    (item[1].transaction.title or "").casefold(),
                    item[1].source_path.casefold(),
                ),
            ):
                key = (statement_index, transaction_index)
                transaction = occurrence.transaction
                tree.insert(
                    "",
                    tk.END,
                    iid=f"{statement_index}-{transaction_index}",
                    values=(
                        transaction.date.strftime("%Y-%m-%d"),
                        self._format_amount(transaction.amount),
                        transaction.title or transaction.description,
                        os.path.basename(occurrence.source_path),
                        "Excluded" if key in exclusions else "Included",
                    ),
                )

        def update_selected(exclude):
            for item in tree.selection():
                statement_index, transaction_index = map(int, item.split("-", 1))
                key = (statement_index, transaction_index)
                if exclude:
                    exclusions.add(key)
                else:
                    exclusions.discard(key)
            on_change()
            refresh()

        def keep_first_copy():
            for group in groups:
                first_source_key = group.occurrences[0].source_key
                for occurrence in group.occurrences:
                    key = (
                        occurrence.statement_index,
                        occurrence.transaction_index,
                    )
                    if occurrence.source_key == first_source_key:
                        exclusions.discard(key)
                    else:
                        exclusions.add(key)
            on_change()
            refresh()

        def restore_all():
            exclusions.clear()
            on_change()
            refresh()

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        RoundedButton(
            buttons,
            text="Exclude Selected",
            command=lambda: update_selected(True),
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons,
            text="Restore Selected",
            command=lambda: update_selected(False),
        ).pack(side=tk.LEFT, padx=(8, 0))
        RoundedButton(
            buttons,
            text="Keep First Copy",
            command=keep_first_copy,
        ).pack(side=tk.LEFT, padx=(8, 0))
        RoundedButton(
            buttons,
            text="Restore All",
            command=restore_all,
        ).pack(side=tk.LEFT, padx=(8, 0))
        RoundedButton(
            buttons, text="Done", style="Primary.TButton", command=dialog.destroy
        ).pack(side=tk.RIGHT)
        refresh()
        self._center_dialog(dialog)
        dialog.grab_set()
        self.wait_window(dialog)

    def _refresh_monthly_duplicate_state(self):
        self.monthly_duplicate_groups = find_duplicate_groups(
            self.statements, self.selected_files
        )
        policy = self.parser.duplicate_policy
        self.monthly_effective_exclusions = set()
        if policy == "warn_review":
            self.monthly_effective_exclusions.update(
                self.monthly_duplicate_exclusions
            )
        elif policy == "auto_merge":
            self.monthly_effective_exclusions.update(
                self.monthly_duplicate_exclusions
            )
            self.monthly_effective_exclusions.update(
                duplicate_occurrences_to_skip(self.monthly_duplicate_groups)
            )
        self.monthly_analysis_statements = filter_duplicate_occurrences(
            self.statements, self.monthly_effective_exclusions
        )
        duplicate_count = len(self.monthly_duplicate_groups)
        self.monthly_review_button.configure(
            text=f"Review Duplicates ({duplicate_count})",
            state=tk.NORMAL
            if duplicate_count and policy == "warn_review"
            else tk.DISABLED,
        )

    def review_monthly_duplicates(self):
        self._show_duplicate_review(
            self.statements,
            self.selected_files,
            self.monthly_duplicate_exclusions,
            self._refresh_preview,
            "Monthly Duplicate Review",
        )
        self.monthly_duplicates_reviewed = True
        self._refresh_preview()

    def export_history_data(self):
        if self.history_analysis is None:
            return
        output_path = GenerateOutput.default_history_output_path(
            self.history_analysis.months
        )
        if os.path.exists(output_path) and not messagebox.askyesno(
            "Replace Historical Workbook?",
            f"Replace the existing workbook?\n\n{output_path}",
            parent=self,
        ):
            return
        try:
            GenerateOutput.create_history_excel(
                self.history_analysis_statements,
                self.history_files,
                self.history_analysis,
                output_path,
                self.history_duplicate_groups,
                self.history_effective_exclusions,
            )
        except (OSError, ValueError) as error:
            messagebox.showerror(
                "History Export Failed",
                f"Could not create the historical workbook:\n{error}",
                parent=self,
            )
            return
        self.status_var.set(f"Historical workbook created: {output_path}")
        messagebox.showinfo(
            "Export Complete",
            f"Your historical analysis is ready:\n{output_path}",
            parent=self,
        )

    def _update_transaction_filter_options(self):
        categories = ("All Categories", "Needs Review", *self.parser.categories)
        sources = {"All Sources"}
        months = set()
        for statement in self.monthly_analysis_statements:
            for transaction in statement.transactions:
                sources.add(transaction.bank.value)
                months.add(transaction.date.strftime("%Y-%m"))
        self.transaction_category_filter.configure(values=categories)
        self.transaction_source_filter.configure(
            values=("All Sources", *sorted(sources - {"All Sources"}))
        )
        month_values = ("All Months", *sorted(months, reverse=True))
        self.transaction_month_filter.configure(values=month_values)
        if self.transaction_category_filter_var.get() not in categories:
            self.transaction_category_filter_var.set("All Categories")
        if self.transaction_source_filter_var.get() not in sources:
            self.transaction_source_filter_var.set("All Sources")
        if self.transaction_month_filter_var.get() not in month_values:
            self.transaction_month_filter_var.set("All Months")

    def _update_overview_month_options(self):
        months = sorted(
            {
                transaction.date.strftime("%Y-%m")
                for statement in self.monthly_analysis_statements
                for transaction in statement.transactions
            },
            reverse=True,
        )
        self.overview_month_combo.configure(values=months)
        if self.overview_month_var.get() not in months:
            self.overview_month_var.set(months[0] if months else "")
        self._refresh_monthly_overview()

    def _refresh_monthly_overview(self):
        month = self.overview_month_var.get()
        self.overview_message_var.set(
            f"Showing activity for {month} against your saved monthly targets."
            if month
            else "Load CSV files to see your monthly budget overview."
        )
        targets = self.target_handler.targets
        rows, total_spent, total_budget, total_income = self._monthly_overview_data(
            self.monthly_analysis_statements, targets, month
        )
        remaining = total_budget - total_spent
        net_income = total_income - total_spent
        self.overview_spent_var.set(self._format_amount(total_spent))
        self.overview_budget_var.set(self._format_amount(total_budget))
        self.overview_remaining_var.set(
            f"{self._format_amount(remaining)} left"
            if remaining >= 0
            else f"Over by {self._format_amount(abs(remaining))}"
        )
        self.overview_income_var.set(self._format_amount(total_income))
        self.overview_net_income_var.set(self._format_amount(net_income))
        self._set_monthly_metric_accent(
            "remaining", STATUS_UNDER if remaining >= 0 else STATUS_OVER
        )
        self._set_monthly_metric_accent(
            "net", STATUS_UNDER if net_income >= 0 else STATUS_OVER
        )
        for widget in self.overview_rows_frame.winfo_children():
            widget.destroy()
        self._overview_table_labels = list(self._overview_header_labels)
        self._overview_column_containers = [self._overview_header_frame]
        for index, row in enumerate(rows):
            category, spent, target, amount_remaining = row
            if target <= 0:
                status = "No target set"
                status_style = "NoTarget"
            elif amount_remaining >= 0:
                status = f"{self._format_amount(amount_remaining)} left"
                status_style = "Under"
            else:
                status = f"Over by {self._format_amount(abs(amount_remaining))}"
                status_style = "Over"
            stripe = "Even" if index % 2 == 0 else "Odd"
            row_frame = ttk.Frame(
                self.overview_rows_frame,
                style=f"DataTable.{stripe}.TFrame",
            )
            row_frame.grid(
                row=index * 2,
                column=0,
                sticky="ew",
            )
            self._overview_column_containers.append(row_frame)
            for column, weight in enumerate(self._overview_column_weights):
                row_frame.columnconfigure(
                    column, weight=weight, uniform="monthly-overview"
                )
            cells = (
                category,
                self._format_amount(spent),
                self._format_amount(target),
                status,
            )
            for column, value in enumerate(cells):
                style = (
                    f"DataTable.{stripe}.{status_style}.TLabel"
                    if column in (1, 3)
                    else f"DataTable.{stripe}.TLabel"
                )
                anchor = ("w", "e", "e", "center")[column]
                label = ttk.Label(
                    row_frame,
                    text=value,
                    style=style,
                    anchor=anchor,
                    width=1,
                    wraplength=1,
                )
                label.grid(row=0, column=column, sticky="ew")
                self._overview_table_labels.append((label, column))
            ttk.Separator(
                self.overview_rows_frame,
                orient=tk.HORIZONTAL,
                style="DataTable.Horizontal.TSeparator",
            ).grid(row=index * 2 + 1, column=0, sticky="ew")
        self.overview_rows_frame.update_idletasks()
        self._update_overview_label_wraps(self.overview_canvas.winfo_width())
        self.overview_rows_frame.update_idletasks()
        self._update_overview_scroll_region()
        self.overview_canvas.yview_moveto(0)

    def _update_overview_scroll_region(self, _event=None):
        self.overview_canvas.configure(
            scrollregion=self.overview_canvas.bbox("all")
        )

    def _resize_overview_rows(self, event):
        self.overview_canvas.itemconfigure(
            self._overview_rows_window, width=event.width
        )
        self._update_overview_label_wraps(event.width)

    def _update_overview_label_wraps(self, width):
        total_weight = sum(self._overview_column_weights)
        for label, column in self._overview_table_labels:
            column_width = width * self._overview_column_weights[column] / total_weight
            label.configure(wraplength=max(1, int(column_width) - 24))

    def _position_overview_resize_handles(self, _event=None):
        header_width = self._overview_header_frame.winfo_width()
        header_height = self._overview_header_frame.winfo_height()
        column_width = max(
            1, header_width - self._overview_header_spacer.winfo_width()
        )
        total_weight = sum(self._overview_column_weights)
        cumulative_weight = 0
        for index, handle in enumerate(self._overview_column_resize_handles):
            cumulative_weight += self._overview_column_weights[index]
            x_position = round(column_width * cumulative_weight / total_weight)
            handle.place(
                x=x_position - 2,
                y=0,
                width=4,
                height=header_height,
            )

    def _start_overview_column_resize(self, boundary, event):
        column_width = max(
            1,
            self._overview_header_frame.winfo_width()
            - self._overview_header_spacer.winfo_width(),
        )
        self._overview_resize_drag = (
            boundary,
            event.x_root,
            self._overview_column_weights.copy(),
            column_width,
        )
        event.widget.configure(background=ACCENT)

    def _drag_overview_column_resize(self, boundary, event):
        if self._overview_resize_drag is None:
            return
        drag_boundary, start_x, original_weights, table_width = (
            self._overview_resize_drag
        )
        if boundary != drag_boundary:
            return

        weights = original_weights.copy()
        total_weight = sum(weights)
        delta_weight = round((event.x_root - start_x) * total_weight / table_width)
        pair_weight = weights[boundary] + weights[boundary + 1]
        minimum_width_pair = (
            self._overview_minimum_column_widths[boundary]
            + self._overview_minimum_column_widths[boundary + 1]
        )
        pair_width = pair_weight * table_width / total_weight
        if pair_width < minimum_width_pair:
            minimum_left = max(
                1,
                round(
                    pair_weight
                    * self._overview_minimum_column_widths[boundary]
                    / minimum_width_pair
                ),
            )
            minimum_right = pair_weight - minimum_left
        else:
            minimum_left = (
                self._overview_minimum_column_widths[boundary] * total_weight
                + table_width
                - 1
            ) // table_width
            minimum_right = (
                self._overview_minimum_column_widths[boundary + 1] * total_weight
                + table_width
                - 1
            ) // table_width
        lower_bound = minimum_left
        upper_bound = pair_weight - minimum_right
        if lower_bound > upper_bound:
            return

        weights[boundary] = min(
            upper_bound,
            max(lower_bound, weights[boundary] + delta_weight),
        )
        weights[boundary + 1] = pair_weight - weights[boundary]
        self._overview_column_weights = weights
        self._apply_overview_column_weights()

    def _finish_overview_column_resize(self, event):
        self._overview_resize_drag = None
        event.widget.configure(background=BORDER)

    def _apply_overview_column_weights(self):
        for container in self._overview_column_containers:
            for column, weight in enumerate(self._overview_column_weights):
                container.columnconfigure(
                    column,
                    weight=weight,
                    uniform="monthly-overview",
                )
        self.overview_rows_frame.update_idletasks()
        self._update_overview_label_wraps(self.overview_canvas.winfo_width())
        self._position_overview_resize_handles()

    def _scroll_overview_rows(self, event):
        widget = event.widget
        while widget is not None and widget is not self.overview_canvas:
            widget = getattr(widget, "master", None)
        if widget is None:
            return
        delta = getattr(event, "delta", 0)
        if delta:
            units = -int(delta / 120) if abs(delta) >= 120 else -int(delta)
            if units == 0:
                units = -1 if delta > 0 else 1
        else:
            units = -1 if getattr(event, "num", None) == 4 else 1
        self.overview_canvas.yview_scroll(units, "units")
        return "break"

    @staticmethod
    def _monthly_overview_data(statements, targets, month):
        category_totals = {}
        income_total = 0.0
        for statement in statements:
            for transaction in statement.transactions:
                if transaction.date.strftime("%Y-%m") != month:
                    continue
                if transaction.is_income():
                    income_total += transaction.amount
                elif transaction.is_purchase():
                    category = category_name(transaction.category)
                    category_totals[category] = (
                        category_totals.get(category, 0.0) + transaction.amount
                    )

        spent_by_category = {
            category: abs(amount) for category, amount in category_totals.items()
        }

        categories = (set(targets) | set(spent_by_category)) - {
            Category.INCOME.value,
            Category.IGNORE.value,
        }
        rows = []
        for category in sorted(categories, key=str.casefold):
            spent = spent_by_category.get(category, 0.0)
            target = targets.get(category, 0.0)
            rows.append((category, spent, target, target - spent))
        total_spent = sum(row[1] for row in rows)
        total_budget = sum(row[2] for row in rows)
        return rows, total_spent, total_budget, abs(income_total)

    def _clear_transaction_filters(self):
        self.transaction_search_var.set("")
        self.transaction_category_filter_var.set("All Categories")
        self.transaction_source_filter_var.set("All Sources")
        self.transaction_month_filter_var.set("All Months")
        self._update_transaction_filter_button()
        self._refresh_transaction_table()

    def choose_files(self, parent=None):
        paths = filedialog.askopenfilenames(
            title="Choose Bank Statement CSV Files",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
            parent=self if parent is None else parent,
        )
        if not paths:
            return

        self._import_files(paths)

    def manage_files(self):
        dialog = tk.Toplevel(self)
        dialog.configure(background=BACKGROUND)
        dialog.title("Manage Statement Files")
        dialog.transient(self)
        dialog.geometry("900x600")
        dialog.minsize(680, 480)
        frame = ttk.Frame(dialog, style="App.TFrame", padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text="Manage Statement Files", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text="Add more statements or remove files from this import. Source files on disk are not changed.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(5, 14))
        file_tree = self._create_treeview(
            frame,
            columns=("path",),
            show="headings",
            selectmode="extended",
            height=6,
        )
        file_tree.heading("path", text="Selected CSV Files")
        file_tree.column("path", width=820, minwidth=300)
        file_scroll = ttk.Scrollbar(
            frame, orient=tk.VERTICAL, command=file_tree.yview
        )
        file_tree.configure(yscrollcommand=file_scroll.set)
        file_tree.grid(row=2, column=0, sticky="nsew")
        file_scroll.grid(row=2, column=1, sticky="ns")

        def refresh_files():
            for item in file_tree.get_children():
                file_tree.delete(item)
            for index, path in enumerate(self.selected_files):
                file_tree.insert("", tk.END, iid=f"file-{index}", values=(path,))

        def add_files():
            self.choose_files(parent=dialog)
            refresh_files()

        def remove_files():
            selection = file_tree.selection()
            if not selection:
                messagebox.showinfo(
                    "Select Files",
                    "Select one or more files to remove from this import.",
                    parent=dialog,
                )
                return
            paths_to_remove = {
                file_tree.item(item, "values")[0] for item in selection
            }
            if not messagebox.askyesno(
                "Remove Selected Files?",
                "Remove the selected files and their transactions from this import? "
                "The original CSV files will not be deleted.",
                parent=dialog,
            ):
                return
            remaining_paths = [
                path for path in self.selected_files if path not in paths_to_remove
            ]
            self._import_files(remaining_paths, replace=True)
            refresh_files()

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        RoundedButton(
            buttons,
            text="Add CSV Files",
            style="Primary.TButton",
            command=add_files,
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons,
            text="Remove Selected",
            command=remove_files,
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(
            buttons, text="Done", command=dialog.destroy
        ).pack(side=tk.RIGHT)
        refresh_files()
        self._center_dialog(dialog)

    def _handle_file_drop(self, event):
        try:
            paths = self.tk.splitlist(event.data)
        except tk.TclError as error:
            messagebox.showerror(
                "Could Not Read Dropped Files", str(error), parent=self
            )
            return COPY

        selected_page = self._selected_workspace_page
        if selected_page == self.history_page:
            workbook_paths = [
                path
                for path in paths
                if path.lower().endswith((".xlsx", ".xlsm"))
            ]
            ignored_paths = [path for path in paths if path not in workbook_paths]
            if not workbook_paths:
                messagebox.showinfo(
                    "Drop Excel Workbooks",
                    "Drop one or more CombinedStatement Excel workbooks into History & Trends.",
                    parent=self,
                )
                return COPY
            if ignored_paths:
                ignored_names = "\n".join(
                    os.path.basename(path) for path in ignored_paths
                )
                messagebox.showwarning(
                    "Unsupported Files Ignored",
                    "Only Excel workbooks can be added to History & Trends. "
                    f"These files were ignored:\n\n{ignored_names}",
                    parent=self,
                )
            self._add_history_files(workbook_paths)
            return COPY

        csv_paths = [path for path in paths if path.lower().endswith(".csv")]
        ignored_paths = [path for path in paths if not path.lower().endswith(".csv")]
        if not csv_paths:
            messagebox.showinfo(
                "Drop CSV Files",
                "Drop one or more CSV statement files into the app.",
                parent=self,
            )
            return COPY
        if ignored_paths:
            ignored_names = "\n".join(os.path.basename(path) for path in ignored_paths)
            messagebox.showwarning(
                "Non-CSV Files Ignored",
                f"Only CSV files can be imported. These files were ignored:\n\n{ignored_names}",
                parent=self,
            )

        self._import_files(csv_paths)
        return COPY

    @staticmethod
    def _file_identity(path):
        return os.path.normcase(os.path.realpath(os.path.expanduser(path)))

    def _import_files(self, paths, replace=False):
        requested_paths = list(paths)
        if replace and not requested_paths:
            self.statements = []
            self.selected_files = []
            self.monthly_duplicate_exclusions.clear()
            self.monthly_duplicates_reviewed = False
            self._refresh_preview()
            self.status_var.set("All statement files were removed from this import.")
            return

        existing_paths = [] if replace else self.selected_files
        seen_paths = {self._file_identity(path) for path in existing_paths}
        paths_to_import = []
        duplicate_paths = []
        for path in requested_paths:
            identity = self._file_identity(path)
            if identity in seen_paths:
                duplicate_paths.append(path)
                continue
            seen_paths.add(identity)
            paths_to_import.append(path)

        if not paths_to_import:
            if duplicate_paths:
                messagebox.showinfo(
                    "Files Already Loaded",
                    "Those CSV files are already included in this import.",
                    parent=self,
                )
            return

        statements = []
        imported_paths = []
        errors = []
        for path in paths_to_import:
            try:
                statements.append(self.parser.statement_from_file(None, path))
                imported_paths.append(path)
            except (OSError, ValueError) as error:
                errors.append(f"{os.path.basename(path)}: {error}")

        if not statements:
            messagebox.showerror(
                "Could Not Import Statements",
                "\n\n".join(errors) or "No transactions could be read from those files.",
                parent=self,
            )
            return

        if replace:
            self.statements = statements
            self.selected_files = imported_paths
            self.monthly_duplicate_exclusions.clear()
        else:
            self.statements.extend(statements)
            self.selected_files.extend(imported_paths)
        self.monthly_duplicates_reviewed = False
        self._refresh_preview()

        if errors:
            messagebox.showwarning(
                "Some Files Were Skipped",
                "\n\n".join(errors),
                parent=self,
            )
        if duplicate_paths:
            messagebox.showinfo(
                "Already Loaded",
                "Files already in this import were skipped:\n\n"
                + "\n".join(os.path.basename(path) for path in duplicate_paths),
                parent=self,
            )

    def _refresh_preview(self):
        self._refresh_monthly_duplicate_state()
        self._update_transaction_filter_options()
        self._update_overview_month_options()
        self.review_groups = {}
        for item in self.review_tree.get_children():
            self.review_tree.delete(item)

        for statement_index, statement in enumerate(self.statements):
            for transaction_index, transaction in enumerate(statement.transactions):
                if (
                    statement_index,
                    transaction_index,
                ) in self.monthly_effective_exclusions:
                    continue
                if transaction.category == Category.NONE:
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
            banks = ", ".join(
                sorted({transaction.bank.value for transaction in transactions})
            )
            self.review_tree.insert(
                "",
                tk.END,
                iid=f"merchant-{index}",
                values=(
                    label,
                    banks,
                    len(transactions),
                    self._format_amount(total),
                ),
            )
            self.review_tree.set(f"merchant-{index}", "merchant", label)
            self.review_tree.item(
                f"merchant-{index}", tags=(key,)
            )

        self._refresh_transaction_table()

    def _refresh_transaction_table(self, _event=None):
        for item in self.transaction_tree.get_children():
            self.transaction_tree.delete(item)
        transaction_count = sum(
            len(statement.transactions)
            for statement in self.monthly_analysis_statements
        )
        visible_transaction_count = 0
        filters = (
            self.transaction_search_var.get(),
            self.transaction_category_filter_var.get(),
            self.transaction_source_filter_var.get(),
            self.transaction_month_filter_var.get(),
        )
        for statement_index, statement in enumerate(self.statements):
            for transaction_index, transaction in enumerate(statement.transactions):
                if (
                    statement_index,
                    transaction_index,
                ) in self.monthly_effective_exclusions:
                    continue
                category = transaction.category
                category_label = (
                    "Needs review"
                    if category == Category.NONE
                    else category_name(category)
                )
                if not self._matches_transaction_filters(
                    transaction,
                    category_label,
                    *filters,
                ):
                    continue
                visible_transaction_count += 1
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

        needs_review = sum(
            len(transactions) for transactions in self.review_groups.values()
        )
        self.export_button.configure(
            state=tk.NORMAL if transaction_count and not needs_review else tk.DISABLED
        )
        self._update_file_summary()
        self.summary_var.set(
            f"Showing {visible_transaction_count} of {transaction_count} "
            f"transaction{'s' if transaction_count != 1 else ''} "
            f"- {needs_review} need categorizing"
            + (
                f" - {len(self.monthly_duplicate_groups)} potential duplicate group(s)"
                if self.parser.duplicate_policy == "warn_review"
                and self.monthly_duplicate_groups
                else ""
            )
        )
        self.status_var.set(
            "Categorize every merchant group before exporting."
            if needs_review
            else "All transactions are categorized. You can export your spreadsheet."
            if transaction_count
            else "No transactions found in the selected files."
        )

    @staticmethod
    def _matches_transaction_filters(
        transaction,
        category_label,
        search,
        selected_category,
        selected_source,
        selected_month,
    ):
        search = search.strip().casefold()
        if search:
            searchable_text = " ".join(
                (
                    transaction.date.strftime("%Y-%m-%d"),
                    transaction.description,
                    transaction.title,
                    category_label,
                    transaction.bank.value,
                )
            ).casefold()
            if search not in searchable_text:
                return False

        if selected_category == "Needs Review":
            if transaction.category != Category.NONE:
                return False
        elif selected_category != "All Categories":
            if category_name(transaction.category) != selected_category:
                return False

        if (
            selected_source != "All Sources"
            and transaction.bank.value != selected_source
        ):
            return False

        return (
            selected_month == "All Months"
            or transaction.date.strftime("%Y-%m") == selected_month
        )

    def _selected_group_keys(self) -> list[str]:
        keys = []
        for item in self.review_tree.selection():
            tags = self.review_tree.item(item, "tags")
            if tags:
                keys.append(tags[0])
        return keys

    def _select_review_group(self, _event=None):
        self.title_var.set("")
        self.rule_pattern_var.set("")
        self.rule_match_type_var.set("Contains")
        self.category_var.set("")

    def apply_categories(self):
        keys = self._selected_group_keys()
        selected_category = self.category_var.get()
        category = category_from_name(selected_category) if selected_category else None
        if not keys:
            messagebox.showinfo(
                "Select Merchants", "Select one or more merchants to categorize.", parent=self
            )
            return
        if category is None or category_name(category) == Category.NONE.value:
            messagebox.showinfo(
                "Choose a Category", "Choose a category before applying it.", parent=self
            )
            return

        title = self.title_var.get().strip()
        pattern = self.rule_pattern_var.get().strip()
        match_type = self.rule_match_type_var.get().casefold()
        if pattern and any(
            not self.parser.category_rule_matches(
                pattern,
                transaction.description,
                match_type,
            )
            for key in keys
            for transaction in self.review_groups[key]
        ):
            messagebox.showinfo(
                "Rule Does Not Match Selected Merchants",
                "The match text must match every selected merchant description. "
                "Shorten it or select only matching merchants.",
                parent=self,
            )
            return
        rules = {}
        for key in keys:
            transactions = self.review_groups[key]
            rule_pattern = pattern or transactions[0].description
            display_title = title or pattern or transactions[0].description
            rules[rule_pattern] = (display_title, category)
        try:
            self.parser.save_category_rules(rules, match_type=match_type)
        except (OSError, sqlite3.Error) as error:
            messagebox.showerror(
                "Could Not Save Merchant Rules",
                f"Your categories were not applied because the rules could not be saved:\n{error}",
                parent=self,
            )
            return
        except ValueError as error:
            messagebox.showerror("Could Not Categorize Merchants", str(error), parent=self            )
            return

        self.parser.apply_category_rules(self.statements)
        self._refresh_preview()
        self.status_var.set(
            f"Categorized {len(keys)} merchant group{'s' if len(keys) != 1 else ''} "
            "and applied matching rules to loaded transactions."
        )

    def export_data(self):
        if not self.statements:
            messagebox.showinfo(
                "No Statements Loaded", "Choose statement CSV files before exporting.", parent=self
            )
            return
        if (
            self.monthly_duplicate_groups
            and self.parser.duplicate_policy == "warn_review"
            and not self.monthly_duplicates_reviewed
        ):
            review_now = messagebox.askyesnocancel(
                "Potential Duplicates",
                "Potential duplicate transactions were found. Review them before "
                "exporting? Choose No to keep all matching transactions.",
                parent=self,
            )
            if review_now is None:
                return
            if review_now:
                self.review_monthly_duplicates()
            else:
                self.monthly_duplicates_reviewed = True
        report_statements = self.monthly_analysis_statements
        if not any(statement.transactions for statement in report_statements):
            messagebox.showinfo(
                "No Transactions to Export",
                "The current duplicate settings exclude all transactions.",
                parent=self,
            )
            return
        if any(
            transaction.category == Category.NONE
            for statement in report_statements
            for transaction in statement.transactions
        ):
            messagebox.showinfo(
                "Review Categories First",
                "Categorize the remaining merchant groups before exporting so no transactions are left out.",
                parent=self,
            )
            return
        output_path = GenerateOutput.default_output_path(report_statements)
        source_files = [
            (statement.bank, path)
            for statement, path in zip(report_statements, self.selected_files)
        ]
        try:
            GenerateOutput.create_excel(
                report_statements, output_path, source_files=source_files
            )
        except (OSError, ValueError) as error:
            messagebox.showerror(
                "Export Failed", f"Could not create the Excel file:\n{error}", parent=self
            )
            return
        self.status_var.set(f"Excel file created: {output_path}")
        messagebox.showinfo(
            "Export Complete", f"Your spreadsheet is ready:\n{output_path}", parent=self
        )

    def _center_dialog(self, dialog: tk.Toplevel, parent=None) -> None:
        parent = self if parent is None else parent
        parent.update_idletasks()
        dialog.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - dialog.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.lift(parent)
        dialog.focus_set()

    def edit_category_rules(self):
        dialog = tk.Toplevel(self)
        dialog.title("Category Rules")
        dialog.configure(background=BACKGROUND)
        dialog.geometry("1120x700")
        dialog.minsize(850, 500)
        dialog.transient(self)

        frame = ttk.Frame(dialog, style="App.TFrame", padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text="Saved Category Rules", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text="Review, edit, or add rules that match transaction descriptions.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(5, 16))

        search_frame = ttk.Frame(frame, style="App.TFrame")
        search_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="Search Rules", style="Subtitle.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        search_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=search_var).grid(
            row=0, column=1, sticky="ew", padx=(0, 8)
        )
        RoundedButton(
            search_frame, text="Clear", command=lambda: search_var.set("")
        ).grid(row=0, column=2)

        columns = ("pattern", "match", "title", "category")
        rules_tree = self._create_treeview(
            frame, columns=columns, show="headings", selectmode="browse"
        )
        heading_labels = {
            "pattern": "Match Text",
            "match": "Match Type",
            "title": "Display Title",
            "category": "Category",
        }
        sort_state = {"column": "pattern", "reverse": False}
        for column, heading, width in (
            ("pattern", heading_labels["pattern"], 300),
            ("match", heading_labels["match"], 140),
            ("title", heading_labels["title"], 340),
            ("category", heading_labels["category"], 190),
        ):
            rules_tree.heading(
                column,
                text=heading,
                command=lambda selected=column: sort_rules(selected),
            )
            rules_tree.column(column, width=width, minwidth=100)
        rules_scroll = ttk.Scrollbar(
            frame, orient=tk.VERTICAL, command=rules_tree.yview
        )
        rules_tree.configure(yscrollcommand=rules_scroll.set)
        rules_tree.grid(row=3, column=0, sticky="nsew")
        rules_scroll.grid(row=3, column=1, sticky="ns")

        def refresh_rules():
            for item in rules_tree.get_children():
                rules_tree.delete(item)
            query = search_var.get().strip().casefold()
            rules = self.parser.category_rules
            if query:
                rules = [
                    rule
                    for rule in rules
                    if query
                    in " ".join(
                        (
                            rule.pattern,
                            rule.match_type,
                            rule.title,
                            category_name(rule.category),
                        )
                    ).casefold()
                ]
            column_index = columns.index(sort_state["column"])
            rules = sorted(
                rules,
                key=lambda rule: (
                    rule.pattern,
                    rule.match_type.title(),
                    rule.title,
                    category_name(rule.category),
                )[column_index].casefold(),
                reverse=sort_state["reverse"],
            )
            for index, rule in enumerate(rules):
                rules_tree.insert(
                    "",
                    tk.END,
                    iid=f"rule-{index}",
                    values=(
                        rule.pattern,
                        rule.match_type.title(),
                        rule.title,
                        category_name(rule.category),
                    ),
                )
            for column, label in heading_labels.items():
                indicator = (
                    " ▼"
                    if sort_state["column"] == column and sort_state["reverse"]
                    else " ▲"
                    if sort_state["column"] == column
                    else ""
                )
                rules_tree.heading(
                    column,
                    text=f"{label}{indicator}",
                    command=lambda selected=column: sort_rules(selected),
                )

        def sort_rules(column):
            if sort_state["column"] == column:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["column"] = column
                sort_state["reverse"] = False
            refresh_rules()

        def close_rule_dialog():
            if self._category_rules_refresh is refresh_rules:
                self._category_rules_refresh = None
            dialog.destroy()

        self._category_rules_refresh = refresh_rules
        dialog.protocol("WM_DELETE_WINDOW", close_rule_dialog)

        def reload_imported_files():
            if self.selected_files:
                self._import_files(self.selected_files, replace=True)

        def open_rule_editor(rule=None):
            editor = tk.Toplevel(dialog)
            editor.title("Edit Category Rule" if rule else "Add Category Rule")
            editor.configure(background=BACKGROUND)
            editor.transient(dialog)
            editor.resizable(False, False)
            fields = ttk.Frame(editor, style="Card.TFrame", padding=24)
            fields.pack(fill=tk.BOTH, expand=True)
            fields.columnconfigure(1, weight=1)

            pattern_var = tk.StringVar(value=rule.pattern if rule else "")
            match_type_var = tk.StringVar(
                value=rule.match_type if rule else "contains"
            )
            title_var = tk.StringVar(value=rule.title if rule else "")
            category_var = tk.StringVar(
                value=category_name(rule.category) if rule else ""
            )

            ttk.Label(fields, text="Match Text", style="CardTitle.TLabel").grid(
                row=0, column=0, sticky="w", padx=(0, 16), pady=8
            )
            pattern_entry = ttk.Entry(fields, textvariable=pattern_var, width=42)
            pattern_entry.grid(row=0, column=1, sticky="ew", pady=8)
            ttk.Label(fields, text="Match Type", style="CardTitle.TLabel").grid(
                row=1, column=0, sticky="w", padx=(0, 16), pady=8
            )
            ttk.Combobox(
                fields,
                textvariable=match_type_var,
                values=("contains", "exact"),
                state="readonly",
            ).grid(row=1, column=1, sticky="ew", pady=8)
            ttk.Label(fields, text="Display Title", style="CardTitle.TLabel").grid(
                row=2, column=0, sticky="w", padx=(0, 16), pady=8
            )
            ttk.Entry(fields, textvariable=title_var).grid(
                row=2, column=1, sticky="ew", pady=8
            )
            ttk.Label(fields, text="Category", style="CardTitle.TLabel").grid(
                row=3, column=0, sticky="w", padx=(0, 16), pady=8
            )
            ttk.Combobox(
                fields,
                textvariable=category_var,
                values=self.parser.categories,
                state="readonly",
            ).grid(row=3, column=1, sticky="ew", pady=8)
            ttk.Label(
                fields,
                text="Contains matches part of a description; exact matches the full normalized description.",
                style="Muted.TLabel",
                wraplength=560,
            ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 12))

            def save_rule():
                category = category_var.get()
                if not category or category == Category.NONE.value:
                    messagebox.showinfo(
                        "Choose a Category",
                        "Choose a category or Ignore before saving.",
                        parent=editor,
                    )
                    return
                try:
                    self.parser.save_category_rule(
                        pattern_var.get(),
                        title_var.get(),
                        category,
                        match_type_var.get(),
                        original_pattern=rule.pattern if rule else None,
                    )
                except (sqlite3.Error, ValueError) as error:
                    messagebox.showerror(
                        "Could Not Save Rule", str(error), parent=editor
                    )
                    return
                editor.destroy()
                refresh_rules()
                self.status_var.set("Category rule saved.")
                reload_imported_files()

            buttons = ttk.Frame(fields, style="Card.TFrame")
            buttons.grid(row=5, column=0, columnspan=2, sticky="e", pady=(8, 0))
            RoundedButton(buttons, text="Cancel", command=editor.destroy).pack(
                side=tk.RIGHT, padx=(8, 0)
            )
            RoundedButton(
                buttons,
                text="Save Rule",
                style="Primary.TButton",
                command=save_rule,
            ).pack(side=tk.RIGHT)
            self._center_dialog(editor, dialog)
            pattern_entry.focus_set()
            editor.grab_set()

        def edit_selected_rule():
            selection = rules_tree.selection()
            if not selection:
                messagebox.showinfo(
                    "Select a Rule", "Select a rule to edit.", parent=dialog
                )
                return
            pattern = rules_tree.item(selection[0], "values")[0]
            rule = next(
                (
                    item
                    for item in self.parser.category_rules
                    if item.pattern == pattern
                ),
                None,
            )
            if rule is not None:
                open_rule_editor(rule)

        def delete_selected_rule():
            selection = rules_tree.selection()
            if not selection:
                messagebox.showinfo(
                    "Select a Rule", "Select a rule to delete.", parent=dialog
                )
                return
            pattern = rules_tree.item(selection[0], "values")[0]
            if not messagebox.askyesno(
                "Delete Category Rule",
                f"Delete the rule for {pattern!r}?",
                parent=dialog,
            ):
                return
            try:
                self.parser.delete_category_rule(pattern)
            except sqlite3.Error as error:
                messagebox.showerror(
                    "Could Not Delete Rule", str(error), parent=dialog
                )
                return
            refresh_rules()
            self.status_var.set("Category rule deleted.")
            reload_imported_files()

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=4, column=0, sticky="ew", pady=(16, 0))
        RoundedButton(
            buttons,
            text="Add Rule",
            style="Primary.TButton",
            command=open_rule_editor,
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons, text="Edit Selected", command=edit_selected_rule
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(
            buttons, text="Delete Selected", command=delete_selected_rule
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(
            buttons,
            text="Done",
            style="Primary.TButton",
            command=close_rule_dialog,
        ).pack(side=tk.RIGHT)
        search_var.trace_add("write", lambda *_: refresh_rules())
        refresh_rules()
        self._center_dialog(dialog)

    def open_settings(self):
        dialog = tk.Toplevel(self)
        dialog.configure(background=BACKGROUND)
        dialog.title("Settings")
        dialog.transient(self)
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, style="App.TFrame", padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Settings", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text="Export or import your categories, targets, and merchant rules.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(5, 20))
        ttk.Label(
            frame,
            text="Duplicate Handling",
            style="App.CardTitle.TLabel",
        ).pack(anchor="w", pady=(0, 6))
        duplicate_policy_labels = (
            "Warn and review",
            "Automatically exclude duplicates",
            "Keep all transactions",
        )
        duplicate_policy_values = {
            "warn_review": duplicate_policy_labels[0],
            "auto_merge": duplicate_policy_labels[1],
            "ignore": duplicate_policy_labels[2],
        }
        self.duplicate_policy_var.set(
            duplicate_policy_values.get(
                self.parser.duplicate_policy, duplicate_policy_labels[0]
            )
        )
        policy_combo = ttk.Combobox(
            frame,
            textvariable=self.duplicate_policy_var,
            values=duplicate_policy_labels,
            state="readonly",
            width=38,
        )
        policy_combo.pack(anchor="w", pady=(0, 8))
        ttk.Label(
            frame,
            text=(
                "Matches use the same date, amount, and place across different files. "
                "Automatic exclusion keeps the first copy; keeping all transactions "
                "turns duplicate filtering off."
            ),
            style="Subtitle.TLabel",
            wraplength=560,
        ).pack(anchor="w", pady=(0, 18))
        policy_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self._save_duplicate_policy(duplicate_policy_values),
        )
        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.pack(fill=tk.X)
        RoundedButton(
            buttons,
            text="Export Settings",
            command=lambda: self.export_settings(dialog),
        ).pack(side=tk.LEFT, padx=(0, 8))
        RoundedButton(
            buttons,
            text="Import Settings",
            command=lambda: self.import_settings(dialog),
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons,
            text="Done",
            style="Primary.TButton",
            command=dialog.destroy,
        ).pack(side=tk.RIGHT)
        self._center_dialog(dialog)

    def _save_duplicate_policy(self, policy_labels):
        selected_label = self.duplicate_policy_var.get()
        policy = next(
            (
                value
                for value, label in policy_labels.items()
                if label == selected_label
            ),
            "warn_review",
        )
        try:
            self.parser.set_duplicate_policy(policy)
        except (OSError, sqlite3.Error, ValueError) as error:
            messagebox.showerror(
                "Could Not Save Duplicate Setting",
                f"The duplicate setting was not changed:\n{error}",
                parent=self,
            )
            self.duplicate_policy_var.set(
                policy_labels.get(self.parser.duplicate_policy, "Warn and review")
            )
            return
        self.monthly_duplicates_reviewed = False
        self.history_duplicates_reviewed = False
        self.history_analysis = None
        self._refresh_preview()
        self._refresh_history()

    def export_settings(self, parent):
        output_directory = GenerateOutput.default_output_directory()
        output_path = os.path.join(output_directory, "StatementParser_Settings.db")
        try:
            os.makedirs(output_directory, exist_ok=True)
        except OSError as error:
            messagebox.showerror(
                "Could Not Export Settings", str(error), parent=parent
            )
            return
        if os.path.exists(output_path) and not messagebox.askyesno(
            "Replace Exported Settings?",
            f"Replace the existing settings export?\n\n{output_path}",
            parent=parent,
        ):
            return
        try:
            self.parser.export_settings(output_path)
        except (OSError, sqlite3.Error, ValueError) as error:
            messagebox.showerror(
                "Could Not Export Settings", str(error), parent=parent
            )
            return
        self.status_var.set(f"Settings exported: {output_path}")
        messagebox.showinfo(
            "Settings Exported",
            f"Your categories, targets, and rules were exported to:\n{output_path}",
            parent=parent,
        )

    def import_settings(self, parent):
        output_directory = GenerateOutput.default_output_directory()
        try:
            os.makedirs(output_directory, exist_ok=True)
        except OSError as error:
            messagebox.showerror(
                "Could Not Import Settings", str(error), parent=parent
            )
            return
        source_path = filedialog.askopenfilename(
            title="Import Statement Parser Settings",
            initialdir=output_directory,
            filetypes=(
                ("Statement Parser settings", "*.db"),
                ("All files", "*.*"),
            ),
            parent=parent,
        )
        if not source_path:
            return

        backup_name = (
            "statement_parser_old_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.db"
        )
        backup_path = os.path.join(output_directory, backup_name)
        if not messagebox.askyesno(
            "Replace Current Settings?",
            "Importing replaces all current categories, targets, and rules. "
            f"A backup will be saved first to:\n{backup_path}\n\nContinue?",
            parent=parent,
        ):
            return
        try:
            self.parser.export_settings(backup_path)
        except (OSError, sqlite3.Error, ValueError) as error:
            messagebox.showerror(
                "Could Not Back Up Settings",
                f"Import was canceled; current settings were not changed.\n\n{error}",
                parent=parent,
            )
            return
        try:
            self.parser.import_settings(source_path)
        except (OSError, sqlite3.Error, ValueError) as error:
            messagebox.showerror(
                "Could Not Import Settings",
                "Current settings were not changed. "
                f"The backup is at:\n{backup_path}\n\n{error}",
                parent=parent,
            )
            return

        self._update_category_choices()
        if self._category_rules_refresh:
            self._category_rules_refresh()
        if self._category_targets_refresh:
            self._category_targets_refresh()
        if self.selected_files:
            self._import_files(self.selected_files, replace=True)
        policy_labels = {
            "warn_review": "Warn and review",
            "auto_merge": "Automatically exclude duplicates",
            "ignore": "Keep all transactions",
        }
        self.duplicate_policy_var.set(
            policy_labels.get(self.parser.duplicate_policy, "Warn and review")
        )
        self.monthly_duplicates_reviewed = False
        self.history_duplicates_reviewed = False
        self.history_analysis = None
        self._refresh_preview()
        self._refresh_history()
        self.status_var.set(
            "Categories, targets, merchant rules, and duplicate handling imported."
        )
        messagebox.showinfo(
            "Settings Imported",
            "Categories, targets, merchant rules, and duplicate handling were replaced.\n\n"
            f"Backup: {backup_path}",
            parent=parent,
        )

    def edit_targets(self):
        dialog = tk.Toplevel(self)
        dialog.configure(background=BACKGROUND)
        dialog.title("Monthly Category Targets")
        dialog.transient(self)
        dialog.geometry("900x680")
        dialog.minsize(720, 480)
        frame = ttk.Frame(dialog, style="App.TFrame", padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text="Monthly Category Targets", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text="Add, edit, or remove categories and their monthly targets.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(5, 16))

        columns = ("category", "target")
        category_tree = self._create_treeview(
            frame, columns=columns, show="headings", selectmode="browse"
        )
        category_tree.heading("category", text="Category")
        category_tree.heading("target", text="Monthly Target", anchor="center")
        category_tree.column("category", width=470, minwidth=180)
        category_tree.column("target", width=230, minwidth=130, anchor="center")
        target_scroll = ttk.Scrollbar(
            frame, orient=tk.VERTICAL, command=category_tree.yview
        )
        category_tree.configure(yscrollcommand=target_scroll.set)
        category_tree.grid(row=2, column=0, sticky="nsew")
        target_scroll.grid(row=2, column=1, sticky="ns")

        def refresh_targets():
            for item in category_tree.get_children():
                category_tree.delete(item)
            for index, category in enumerate(self.target_handler.category_details):
                category_tree.insert(
                    "",
                    tk.END,
                    iid=f"category-{index}",
                    values=(category.name, f"${category.amount:,.2f}"),
                )

        def close_target_dialog():
            if self._category_targets_refresh is refresh_targets:
                self._category_targets_refresh = None
            dialog.destroy()

        self._category_targets_refresh = refresh_targets
        dialog.protocol("WM_DELETE_WINDOW", close_target_dialog)

        def selected_category():
            selection = category_tree.selection()
            if not selection:
                messagebox.showinfo(
                    "Select a Category",
                    "Select a category first.",
                    parent=dialog,
                )
                return None
            name = category_tree.item(selection[0], "values")[0]
            return next(
                (
                    category
                    for category in self.target_handler.category_details
                    if category.name == name
                ),
                None,
            )

        def open_category_editor(category=None):
            editor = tk.Toplevel(dialog)
            editor.title("Edit Category" if category else "Add Category")
            editor.configure(background=BACKGROUND)
            editor.transient(dialog)
            editor.resizable(False, False)
            fields = ttk.Frame(editor, style="Card.TFrame", padding=24)
            fields.pack(fill=tk.BOTH, expand=True)
            fields.columnconfigure(1, weight=1)
            name_var = tk.StringVar(value=category.name if category else "")
            amount_var = tk.StringVar(
                value=str(category.amount) if category else "0.0"
            )

            ttk.Label(fields, text="Category Name", style="CardTitle.TLabel").grid(
                row=0, column=0, sticky="w", padx=(0, 16), pady=8
            )
            name_entry = ttk.Entry(fields, textvariable=name_var, width=32)
            name_entry.grid(row=0, column=1, sticky="ew", pady=8)
            if category and category.is_system:
                name_entry.configure(state=tk.DISABLED)
            ttk.Label(fields, text="Monthly Target", style="CardTitle.TLabel").grid(
                row=1, column=0, sticky="w", padx=(0, 16), pady=8
            )
            amount_entry = ttk.Entry(fields, textvariable=amount_var, width=18)
            amount_entry.grid(row=1, column=1, sticky="ew", pady=8)
            if category and category.is_system:
                ttk.Label(
                    fields,
                    text="System category names cannot be changed.",
                    style="Muted.TLabel",
                ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 8))

            def save_category():
                try:
                    self.parser.save_category(
                        name_var.get(),
                        float(amount_var.get()),
                        original_name=category.name if category else None,
                    )
                except (sqlite3.Error, ValueError) as error:
                    messagebox.showerror(
                        "Could Not Save Category", str(error), parent=editor
                    )
                    return
                editor.destroy()
                refresh_targets()
                if self.history_analysis is not None:
                    self._refresh_history_overview_cards()
                self._update_category_choices()
                if self._category_rules_refresh:
                    self._category_rules_refresh()
                if self.selected_files:
                    self._import_files(self.selected_files, replace=True)
                self.status_var.set("Category and monthly target saved.")

            buttons = ttk.Frame(fields, style="Card.TFrame")
            buttons.grid(row=3, column=0, columnspan=2, sticky="e", pady=(10, 0))
            RoundedButton(
                buttons, text="Cancel", command=editor.destroy
            ).pack(side=tk.RIGHT, padx=(8, 0))
            RoundedButton(
                buttons,
                text="Save Category",
                style="Primary.TButton",
                command=save_category,
            ).pack(side=tk.RIGHT)
            self._center_dialog(editor, dialog)
            if not category or not category.is_system:
                name_entry.focus_set()
            else:
                amount_entry.focus_set()
            editor.grab_set()

        def edit_selected_category():
            category = selected_category()
            if category:
                open_category_editor(category)

        def delete_selected_category():
            category = selected_category()
            if category is None:
                return
            if category.is_system:
                messagebox.showinfo(
                    "Protected Category",
                    f"{category.name} is required by the app and cannot be deleted.",
                    parent=dialog,
                )
                return
            if not messagebox.askyesno(
                "Delete Category",
                f"Deleting {category.name!r} also deletes all of its saved rules. "
                "Continue?",
                parent=dialog,
            ):
                return
            try:
                self.parser.delete_category(category.name)
            except (sqlite3.Error, ValueError) as error:
                messagebox.showerror(
                    "Could Not Delete Category", str(error), parent=dialog
                )
                return
            refresh_targets()
            self._update_category_choices()
            if self._category_rules_refresh:
                self._category_rules_refresh()
            if self.selected_files:
                self._import_files(self.selected_files, replace=True)
            self.status_var.set("Category and its saved rules deleted.")

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        RoundedButton(
            buttons,
            text="Add Category",
            style="Primary.TButton",
            command=open_category_editor,
        ).pack(side=tk.LEFT)
        RoundedButton(
            buttons, text="Edit Selected", command=edit_selected_category
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(
            buttons, text="Delete Selected", command=delete_selected_category
        ).pack(side=tk.LEFT, padx=(10, 0))
        RoundedButton(
            buttons,
            text="Done",
            style="Primary.TButton",
            command=close_target_dialog,
        ).pack(side=tk.RIGHT)
        refresh_targets()
        self._center_dialog(dialog)

    def _update_category_choices(self):
        categories = self.parser.categories
        self.category_combo.configure(values=categories)
        if self.category_var.get() not in categories:
            self.category_var.set("")
        self._update_transaction_filter_options()
        self._refresh_monthly_overview()

    @staticmethod
    def _format_amount(amount: float) -> str:
        return f"-${abs(amount):,.2f}" if amount < 0 else f"${amount:,.2f}"

    def log_status(self, message: str):
        self.status_var.set(message)


def main():
    app = BankParseGUI()
    app.mainloop()
