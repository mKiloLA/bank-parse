import unittest
from datetime import date

from statement_parser.__main__ import BankParseGUI
from statement_parser.source.data.DataTypes import Bank, Category
from statement_parser.source.data.Statement import Statement
from statement_parser.source.data.Transaction import Transaction


class GUIDataTests(unittest.TestCase):
    def test_monthly_overview_data_summarizes_spending_income_and_targets(self):
        statements = [
            Statement(
                Bank.CITI,
                [
                    Transaction(
                        date(2026, 9, 1), -12.0, "Cafe", Bank.CITI,
                        Category.DINING
                    ),
                    Transaction(
                        date(2026, 9, 1), 3.0, "Cafe Refund", Bank.CITI,
                        Category.DINING
                    ),
                    Transaction(
                        date(2026, 9, 2), -50.0, "Pet Shop", Bank.CITI,
                        "Pet Care"
                    ),
                    Transaction(
                        date(2026, 9, 3), 200.0, "Paycheck", Bank.CITI,
                        Category.INCOME
                    ),
                    Transaction(
                        date(2026, 9, 4), -18.0, "Transfer", Bank.CITI,
                        Category.IGNORE
                    ),
                    Transaction(
                        date(2026, 9, 5), -15.0, "Unreviewed", Bank.CITI,
                        Category.NONE
                    ),
                    Transaction(
                        date(2026, 10, 1), -20.0, "Cafe", Bank.CITI,
                        Category.DINING
                    ),
                ],
            )
        ]
        targets = {
            Category.DINING.value: 40.0,
            "Pet Care": 30.0,
            Category.INCOME.value: 100.0,
            Category.IGNORE.value: 99.0,
        }

        rows, total_spent, total_budget, total_income = (
            BankParseGUI._monthly_overview_data(statements, targets, "2026-09")
        )

        self.assertEqual(
            rows,
            [
                ("Dining", 9.0, 40.0, 31.0),
                ("Pet Care", 50.0, 30.0, -20.0),
            ],
        )
        self.assertEqual(total_spent, 59.0)
        self.assertEqual(total_budget, 70.0)
        self.assertEqual(total_income, 200.0)

    def test_transaction_filters_match_search_category_source_and_month(self):
        transaction = Transaction(
            date(2026, 9, 12),
            -12.0,
            "Morning Coffee",
            Bank.CITI,
            Category.DINING,
            "Coffee Shop",
        )

        self.assertTrue(
            BankParseGUI._matches_transaction_filters(
                transaction,
                "Dining",
                "coffee",
                "Dining",
                "Citi",
                "2026-09",
            )
        )
        self.assertFalse(
            BankParseGUI._matches_transaction_filters(
                transaction,
                "Dining",
                "",
                "Needs Review",
                "All Sources",
                "All Months",
            )
        )
        self.assertFalse(
            BankParseGUI._matches_transaction_filters(
                transaction,
                "Dining",
                "",
                "All Categories",
                "Venmo",
                "All Months",
            )
        )


if __name__ == "__main__":
    unittest.main()
