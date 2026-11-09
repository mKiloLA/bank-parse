import os
import tempfile
import unittest
from datetime import date

import openpyxl

from statement_parser.source.GenerateOutput import GenerateOutput
from statement_parser.source.HistoryAnalysis import (
    analyze_history,
    duplicate_occurrences_to_skip,
    filter_duplicate_occurrences,
    find_duplicate_groups,
)
from statement_parser.source.data.DataTypes import Bank, Category
from statement_parser.source.data.Statement import Statement
from statement_parser.source.data.Transaction import Transaction


class HistoryAnalysisTests(unittest.TestCase):
    def test_category_average_and_monthly_change_include_zero_months(self):
        statement = Statement(
            Bank.CITI,
            [
                Transaction(
                    date(2026, 9, 2), -100.0, "Dining", Bank.CITI, Category.DINING
                ),
                Transaction(
                    date(2026, 10, 2), -120.0, "Dining", Bank.CITI, Category.DINING
                ),
                Transaction(
                    date(2026, 11, 2), -140.0, "Dining", Bank.CITI, Category.DINING
                ),
                Transaction(
                    date(2026, 9, 4), -90.0, "Power", Bank.CITI, Category.UTILITIES
                ),
                Transaction(
                    date(2026, 11, 4), -120.0, "Power", Bank.CITI, Category.UTILITIES
                ),
                Transaction(
                    date(2026, 9, 1), 2000.0, "Pay", Bank.CITI, Category.INCOME
                ),
                Transaction(
                    date(2026, 10, 1), -50.0, "Transfer", Bank.CITI, Category.IGNORE
                ),
            ],
        )

        analysis = analyze_history([statement])
        categories = {item.category: item for item in analysis.category_history}

        self.assertEqual(len(analysis.months), 3)
        self.assertEqual(categories["Dining"].average_monthly_spend, 120.0)
        self.assertEqual(categories["Dining"].monthly_change, 20.0)
        self.assertEqual(categories["Utilities"].monthly_spending, (90.0, 0.0, 120.0))
        self.assertEqual(analysis.monthly_income, (2000.0, 0.0, 0.0))
        self.assertEqual(analysis.monthly_spending_change, 35.0)

    def test_category_trend_excludes_months_before_first_spend(self):
        statement = Statement(
            Bank.CITI,
            [
                Transaction(
                    date(2026, 9, 2), -90.0, "Power", Bank.CITI, Category.UTILITIES
                ),
                Transaction(
                    date(2026, 10, 2), -100.0, "Dining", Bank.CITI, Category.DINING
                ),
                Transaction(
                    date(2026, 11, 2), -120.0, "Power", Bank.CITI, Category.UTILITIES
                ),
            ],
        )

        analysis = analyze_history([statement])
        dining = next(
            item for item in analysis.category_history if item.category == "Dining"
        )

        self.assertEqual(dining.monthly_spending, (0.0, 100.0, 0.0))
        self.assertEqual(dining.first_spending_month_index, 1)
        self.assertEqual(dining.average_monthly_spend, 50.0)
        self.assertEqual(dining.monthly_change, -100.0)

    def test_subscription_requires_same_merchant_and_amount_in_three_months(self):
        statement = Statement(
            Bank.CITI,
            [
                Transaction(
                    date(2026, month, 5),
                    -14.99,
                    "Streaming Co",
                    Bank.CITI,
                    Category.ENTERTAINMENT,
                    "STREAMING CO",
                )
                for month in (1, 2, 3)
            ]
            + [
                Transaction(
                    date(2026, month, 7), -8.0, "Gym", Bank.CITI, Category.OTHER, "Gym"
                )
                for month in (1, 2)
            ],
        )

        analysis = analyze_history([statement])

        self.assertEqual(len(analysis.subscriptions), 1)
        self.assertEqual(analysis.subscriptions[0].merchant, "STREAMING CO")
        self.assertEqual(analysis.subscriptions[0].amount, 14.99)
        self.assertEqual(len(analysis.subscriptions[0].months), 3)

    def test_duplicate_detection_is_cross_source_and_removal_skips_later_copy(self):
        first = Statement(
            Bank.CITI,
            [
                Transaction(
                    date(2026, 9, 5),
                    -19.99,
                    "NETFLIX",
                    Bank.CITI,
                    Category.ENTERTAINMENT,
                    "Netflix",
                )
            ],
        )
        second = Statement(
            Bank.VENMO,
            [
                Transaction(
                    date(2026, 9, 5),
                    -19.99,
                    "Netflix ",
                    Bank.VENMO,
                    Category.ENTERTAINMENT,
                    " netflix ",
                ),
                Transaction(
                    date(2026, 9, 6),
                    -19.99,
                    "Netflix",
                    Bank.VENMO,
                    Category.ENTERTAINMENT,
                    "Netflix",
                ),
            ],
        )
        statements = [first, second]
        sources = ["first.xlsx", "second.xlsx"]

        groups = find_duplicate_groups(statements, sources)
        excluded = duplicate_occurrences_to_skip(groups)

        self.assertEqual(len(groups), 1)
        self.assertEqual(excluded, {(1, 0)})

    def test_history_workbook_contains_analysis_tables_charts_and_duplicate_audit(self):
        first = Statement(
            Bank.CITI,
            [
                Transaction(
                    date(2026, 1, 5),
                    -15.0,
                    "Stream Service",
                    Bank.CITI,
                    Category.ENTERTAINMENT,
                    "Stream Service",
                ),
                Transaction(
                    date(2026, 1, 8),
                    -100.0,
                    "Dining",
                    Bank.CITI,
                    Category.DINING,
                    "Dining",
                ),
                Transaction(
                    date(2026, 2, 5),
                    -15.0,
                    "Stream Service",
                    Bank.CITI,
                    Category.ENTERTAINMENT,
                    "Stream Service",
                ),
                Transaction(
                    date(2026, 2, 8),
                    -120.0,
                    "Dining",
                    Bank.CITI,
                    Category.DINING,
                    "Dining",
                ),
                Transaction(
                    date(2026, 3, 5),
                    -15.0,
                    "Stream Service",
                    Bank.CITI,
                    Category.ENTERTAINMENT,
                    "Stream Service",
                ),
                Transaction(
                    date(2026, 3, 8),
                    -140.0,
                    "Dining",
                    Bank.CITI,
                    Category.DINING,
                    "Dining",
                ),
            ],
        )
        second = Statement(
            Bank.VENMO,
            [
                Transaction(
                    date(2026, 1, 5),
                    -15.0,
                    "Stream Service",
                    Bank.VENMO,
                    Category.ENTERTAINMENT,
                    "Stream Service",
                )
            ],
        )
        statements = [first, second]
        sources = ["first.xlsx", "second.xlsx"]
        duplicate_groups = find_duplicate_groups(statements, sources)
        excluded = duplicate_occurrences_to_skip(duplicate_groups)
        filtered = filter_duplicate_occurrences(statements, excluded)
        analysis = analyze_history(filtered)

        with tempfile.TemporaryDirectory() as directory:
            output_path = os.path.join(directory, "history.xlsx")
            GenerateOutput.create_history_excel(
                filtered,
                sources,
                analysis,
                output_path,
                duplicate_groups,
                excluded,
            )
            workbook = openpyxl.load_workbook(output_path, data_only=True)

        self.assertEqual(
            workbook.sheetnames,
            [
                "Overview",
                "Category Trends",
                "Subscriptions",
                "Transactions",
                "Analysis Data",
                "Duplicates",
            ],
        )
        self.assertEqual(workbook.active.title, "Overview")
        self.assertEqual(len(workbook["Overview"]._charts), 1)
        self.assertEqual(len(workbook["Category Trends"]._charts), 1)
        self.assertEqual(workbook["Analysis Data"].sheet_state, "hidden")
        self.assertIn("HistoricalCategoryTrends", workbook["Category Trends"].tables)
        trends = workbook["Category Trends"]
        category_table = trends.tables["HistoricalCategoryTrends"]
        min_column, min_row, max_column, max_row = openpyxl.utils.cell.range_boundaries(
            category_table.ref
        )
        headers = [
            trends.cell(row=min_row, column=column).value
            for column in range(min_column, max_column + 1)
        ]
        category_column = min_column + headers.index("Category")
        total_column = min_column + headers.index("Total Spent")
        category_totals = {
            trends.cell(row=row, column=category_column)
            .value: trends.cell(row=row, column=total_column)
            .value
            for row in range(min_row + 1, max_row + 1)
        }
        self.assertEqual(category_totals["Dining"], 360.0)
        self.assertIn("first spending month", trends["B2"].value)
        monthly_table = workbook["Category Trends"].tables[
            "HistoricalCategoryMonthlySpending"
        ]
        matrix_header_row = int(monthly_table.ref.split(":")[0][1:])
        self.assertGreaterEqual(matrix_header_row, 22)
        self.assertEqual(workbook["Duplicates"]["F6"].value, "Excluded from analysis")
        self.assertEqual(
            workbook["Subscriptions"]["B6"].value,
            "Stream Service",
        )


if __name__ == "__main__":
    unittest.main()
