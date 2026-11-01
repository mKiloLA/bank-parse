import os
import tempfile
import unittest
from datetime import date
from unittest.mock import patch

import openpyxl

from statement_parser.source.GenerateOutput import GenerateOutput
from statement_parser.source.data.DataTypes import Bank, Category
from statement_parser.source.data.Statement import Statement
from statement_parser.source.data.Transaction import Transaction


class TargetHandlerStub:
    @property
    def targets(self):
        return {
            category.value: 0.0
            for category in Category
            if category not in (Category.NONE, Category.IGNORE)
        }

    def get_target(self, category):
        return 0.0


class GenerateOutputTests(unittest.TestCase):
    def test_output_filename_uses_the_most_common_transaction_month(self):
        statements = [
            Statement(
                Bank.VENMO,
                [
                    Transaction(
                        date(2026, 9, 3), -12.0, "Lunch", Bank.VENMO,
                        Category.DINING, "Lunch"
                    ),
                    Transaction(
                        date(2026, 9, 4), -8.0, "Coffee", Bank.VENMO,
                        Category.DINING, "Coffee"
                    ),
                ],
            ),
            Statement(
                Bank.CITI,
                [
                    Transaction(
                        date(2026, 10, 1), -20.0, "Store", Bank.CITI,
                        Category.OTHER, "Store"
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
                            date(2026, 9, 3), -12.0, "Lunch", Bank.VENMO,
                            Category.DINING, "Lunch"
                        )
                    ],
                ),
                Statement(
                    Bank.VENMO,
                    [
                        Transaction(
                            date(2026, 9, 4), -8.0, "Coffee", Bank.VENMO,
                            Category.DINING, "Coffee"
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


if __name__ == "__main__":
    unittest.main()
