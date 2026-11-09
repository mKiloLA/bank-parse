import os
import tempfile
import unittest
from datetime import date
from unittest.mock import patch

import openpyxl

from statement_parser.source.GenerateOutput import GenerateOutput
from statement_parser.source.HistoryAnalysis import load_combined_statement
from statement_parser.source.data.DataTypes import Bank, Category
from statement_parser.source.data.Statement import Statement
from statement_parser.source.data.Transaction import Transaction


class TargetHandlerStub:
    def __init__(self, extra_targets=None):
        self._targets = {
            category.value: 0.0
            for category in Category
            if category not in (Category.NONE, Category.IGNORE)
        }
        self._targets.update(extra_targets or {})

    @property
    def targets(self):
        return self._targets

    def get_target(self, category):
        return self._targets.get(category, 0.0)


class GenerateOutputTests(unittest.TestCase):
    def test_output_filename_uses_the_most_common_transaction_month(self):
        statements = [
            Statement(
                Bank.VENMO,
                [
                    Transaction(
                        date(2026, 9, 3),
                        -12.0,
                        "Lunch",
                        Bank.VENMO,
                        Category.DINING,
                        "Lunch",
                    ),
                    Transaction(
                        date(2026, 9, 4),
                        -8.0,
                        "Coffee",
                        Bank.VENMO,
                        Category.DINING,
                        "Coffee",
                    ),
                ],
            ),
            Statement(
                Bank.CITI,
                [
                    Transaction(
                        date(2026, 10, 1),
                        -20.0,
                        "Store",
                        Bank.CITI,
                        Category.OTHER,
                        "Store",
                    )
                ],
            ),
        ]

        self.assertEqual(
            GenerateOutput.output_filename(statements),
            "CombinedStatement_2026-09.xlsx",
        )
        self.assertEqual(
            GenerateOutput.default_output_path(statements),
            os.path.join(
                os.path.expanduser("~/StatementParser"),
                "CombinedStatement_2026-09.xlsx",
            ),
        )

    def test_workbook_includes_raw_csv_sheets_named_for_each_bank(self):
        with tempfile.TemporaryDirectory() as directory:
            first_csv = os.path.join(directory, "venmo-one.csv")
            second_csv = os.path.join(directory, "venmo-two.csv")
            output_path = os.path.join(directory, "combined.xlsx")
            with open(first_csv, "w", encoding="utf-8") as file:
                file.write("Date,Description,Amount\n2026-09-03,=1+1,-12.00\n")
            with open(second_csv, "w", encoding="utf-8") as file:
                file.write("Date,Description,Amount\n2026-09-04,Coffee,-8.00\n")

            statements = [
                Statement(
                    Bank.VENMO,
                    [
                        Transaction(
                            date(2026, 9, 3),
                            -12.0,
                            "Lunch",
                            Bank.VENMO,
                            Category.DINING,
                            "Lunch",
                        )
                    ],
                ),
                Statement(
                    Bank.VENMO,
                    [
                        Transaction(
                            date(2026, 9, 4),
                            -8.0,
                            "Coffee",
                            Bank.VENMO,
                            Category.DINING,
                            "Coffee",
                        )
                    ],
                ),
            ]

            with patch(
                "statement_parser.source.GenerateOutput.CategoryTargetHandler",
                return_value=TargetHandlerStub(),
            ):
                GenerateOutput.create_excel(
                    statements,
                    output_path,
                    source_files=[
                        (Bank.VENMO, first_csv),
                        (Bank.VENMO, second_csv),
                    ],
                )

            workbook = openpyxl.load_workbook(output_path, data_only=False)
            self.assertIn("Venmo", workbook.sheetnames)
            self.assertIn("Venmo (2)", workbook.sheetnames)
            self.assertEqual(workbook["Venmo"]["B2"].value, "=1+1")
            self.assertEqual(workbook["Venmo"]["B2"].data_type, "s")
            self.assertEqual(workbook["Venmo (2)"]["B2"].value, "Coffee")

    def test_workbook_includes_custom_categories(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = os.path.join(directory, "custom-category.xlsx")
            statements = [
                Statement(
                    Bank.NONE,
                    [
                        Transaction(
                            date(2026, 9, 3),
                            -12.0,
                            "Pet Supply Shop",
                            Bank.NONE,
                            "Pet Care",
                            "Pet Supplies",
                        )
                    ],
                )
            ]
            with patch(
                "statement_parser.source.GenerateOutput.CategoryTargetHandler",
                return_value=TargetHandlerStub({"Pet Care": 75.0}),
            ):
                GenerateOutput.create_excel(statements, output_path)

            workbook = openpyxl.load_workbook(output_path, data_only=True)
            self.assertEqual(workbook["Expenses & Income"]["D3"].value, "Pet Care")
            self.assertIn(
                "Pet Care",
                [
                    row[0].value
                    for row in workbook["Summary"].iter_rows(min_col=2, max_col=2)
                ],
            )

    def test_dashboard_summarizes_period_budget_and_adds_charts(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = os.path.join(directory, "dashboard.xlsx")
            statements = [
                Statement(
                    Bank.CITI,
                    [
                        Transaction(
                            date(2026, 9, 3),
                            -100.0,
                            "Rent",
                            Bank.CITI,
                            Category.MORTGAGE,
                            "Rent",
                        ),
                        Transaction(
                            date(2026, 9, 4),
                            2000.0,
                            "Pay",
                            Bank.CITI,
                            Category.INCOME,
                            "Employer",
                        ),
                        Transaction(
                            date(2026, 9, 5),
                            -15.0,
                            "Pet Store",
                            Bank.CITI,
                            "Pet Care",
                            "Pet Store",
                        ),
                    ],
                ),
                Statement(
                    Bank.CITI,
                    [
                        Transaction(
                            date(2026, 10, 3),
                            -50.0,
                            "Rent",
                            Bank.CITI,
                            Category.MORTGAGE,
                            "Rent",
                        ),
                        Transaction(
                            date(2026, 10, 4),
                            2000.0,
                            "Pay",
                            Bank.CITI,
                            Category.INCOME,
                            "Employer",
                        ),
                    ],
                ),
            ]
            with patch(
                "statement_parser.source.GenerateOutput.CategoryTargetHandler",
                return_value=TargetHandlerStub(
                    {
                        Category.MORTGAGE.value: 200.0,
                        Category.INCOME.value: 2000.0,
                        "Pet Care": 75.0,
                    }
                ),
            ):
                GenerateOutput.create_excel(statements, output_path)

            workbook = openpyxl.load_workbook(output_path, data_only=True)
            dashboard = workbook["Dashboard"]
            self.assertEqual(workbook.sheetnames[0], "Dashboard")
            self.assertEqual(dashboard["B6"].value, 4000.0)
            self.assertEqual(dashboard["E6"].value, 165.0)
            self.assertEqual(dashboard["H6"].value, 3835.0)
            self.assertEqual(len(workbook["Dashboard"]._charts), 4)
            self.assertEqual(dashboard.sheet_view.zoomScale, 160)
            self.assertIsNone(dashboard.freeze_panes)
            self.assertEqual(dashboard.sheet_view.selection[0].activeCell, "A1")
            charts = dashboard._charts
            self.assertEqual(charts[0].legend.position, "r")
            self.assertEqual(charts[2].legend.position, "r")
            self.assertGreaterEqual(
                charts[2].anchor._from.row - charts[0].anchor._from.row, 27
            )
            self.assertGreater(charts[3].anchor._from.row, charts[2].anchor._from.row)
            plotting_data = workbook["Plotting Data"]
            self.assertEqual(plotting_data["J1"].value, "Day")
            self.assertEqual(plotting_data["J2"].value, "Sep 03")
            self.assertEqual(plotting_data["K2"].value, 100.0)
            self.assertEqual(plotting_data["J4"].value, "Sep 05")
            self.assertEqual(plotting_data["K4"].value, 15.0)
            category_rows = {
                row[0].value: row
                for row in dashboard.iter_rows(
                    min_row=12, max_row=40, min_col=2, max_col=7
                )
                if row[0].value
            }
            self.assertEqual(category_rows["Pet Care"][2].value, 150.0)
            self.assertEqual(category_rows["Pet Care"][1].value, 15.0)

    def test_history_loader_reads_generated_combined_statement_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = os.path.join(directory, "combined.xlsx")
            statement = Statement(
                Bank.CITI,
                [
                    Transaction(
                        date(2026, 9, 3),
                        -15.0,
                        "Coffee",
                        Bank.CITI,
                        Category.DINING,
                        "Coffee Shop",
                    ),
                    Transaction(
                        date(2026, 9, 1),
                        2000.0,
                        "Pay",
                        Bank.CITI,
                        Category.INCOME,
                        "Employer",
                    ),
                ],
            )
            with patch(
                "statement_parser.source.GenerateOutput.CategoryTargetHandler",
                return_value=TargetHandlerStub(),
            ):
                GenerateOutput.create_excel([statement], output_path)

            loaded = load_combined_statement(output_path)

        self.assertEqual(len(loaded.transactions), 2)
        self.assertEqual(
            {
                (transaction.title, transaction.amount)
                for transaction in loaded.transactions
            },
            {("Coffee Shop", -15.0), ("Employer", 2000.0)},
        )


if __name__ == "__main__":
    unittest.main()
