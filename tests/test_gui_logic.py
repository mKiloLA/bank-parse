import unittest
from datetime import date
from types import SimpleNamespace

from statement_parser.__main__ import BankParseGUI
from statement_parser.source.data.DataTypes import Bank, Category
from statement_parser.source.data.Statement import Statement
from statement_parser.source.data.Transaction import Transaction


class GUIDataTests(unittest.TestCase):
    def test_mouse_wheel_units_handle_platform_event_variants(self):
        events = (
            (SimpleNamespace(delta=120), -1),
            (SimpleNamespace(delta=-240), 2),
            (SimpleNamespace(delta=1), -1),
            (SimpleNamespace(delta=-1), 1),
            (SimpleNamespace(delta=0, num=4), -1),
            (SimpleNamespace(delta=0, num=5), 1),
        )

        for event, expected_units in events:
            with self.subTest(event=event):
                self.assertEqual(BankParseGUI._mouse_wheel_units(event), expected_units)

    def test_mouse_wheel_scroll_routes_nested_widgets_and_respects_history_page(self):
        class ScrollCanvas:
            def __init__(self):
                self.scrolls = []

            def yview_scroll(self, units, mode):
                self.scrolls.append((units, mode))

        canvas = ScrollCanvas()
        nested_widget = SimpleNamespace(master=canvas)
        event = SimpleNamespace(widget=nested_widget, delta=120)
        self.assertEqual(BankParseGUI._scroll_canvas_from_event(event, canvas), "break")
        self.assertEqual(canvas.scrolls, [(-1, "units")])

        outside_event = SimpleNamespace(
            widget=SimpleNamespace(master=None),
            delta=120,
        )
        self.assertIsNone(BankParseGUI._scroll_canvas_from_event(outside_event, canvas))
        self.assertEqual(canvas.scrolls, [(-1, "units")])

        history_page = object()
        controller = SimpleNamespace(
            history_page=history_page,
            _selected_workspace_page=object(),
            history_page_canvas=canvas,
            _scroll_canvas_from_event=BankParseGUI._scroll_canvas_from_event,
        )
        self.assertIsNone(BankParseGUI._scroll_history_page(controller, event))
        controller._selected_workspace_page = history_page
        self.assertEqual(BankParseGUI._scroll_history_page(controller, event), "break")
        self.assertEqual(canvas.scrolls, [(-1, "units"), (-1, "units")])

    def test_monthly_overview_data_summarizes_spending_income_and_targets(self):
        statements = [
            Statement(
                Bank.CITI,
                [
                    Transaction(
                        date(2026, 9, 1), -12.0, "Cafe", Bank.CITI, Category.DINING
                    ),
                    Transaction(
                        date(2026, 9, 1), 3.0, "Cafe Refund", Bank.CITI, Category.DINING
                    ),
                    Transaction(
                        date(2026, 9, 2), -50.0, "Pet Shop", Bank.CITI, "Pet Care"
                    ),
                    Transaction(
                        date(2026, 9, 3), 200.0, "Paycheck", Bank.CITI, Category.INCOME
                    ),
                    Transaction(
                        date(2026, 9, 4), -18.0, "Transfer", Bank.CITI, Category.IGNORE
                    ),
                    Transaction(
                        date(2026, 9, 5), -15.0, "Unreviewed", Bank.CITI, Category.NONE
                    ),
                    Transaction(
                        date(2026, 10, 1), -20.0, "Cafe", Bank.CITI, Category.DINING
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

        (
            rows,
            total_spent,
            total_budget,
            total_income,
        ) = BankParseGUI._monthly_overview_data(statements, targets, "2026-09")

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

    def test_monthly_chart_totals_follow_selected_categories(self):
        categories = (
            ("Dining", 120.0, 200.0, 80.0),
            ("Utilities", 80.0, 100.0, 20.0),
        )

        selected_total, maximum = BankParseGUI._monthly_chart_totals(categories)
        self.assertEqual(selected_total, 200.0)
        self.assertEqual(maximum, 200.0)

        selected_total, maximum = BankParseGUI._monthly_chart_totals((categories[0],))
        self.assertEqual(selected_total, 120.0)
        self.assertEqual(maximum, 120.0)

        selected_total, maximum = BankParseGUI._monthly_chart_totals(())
        self.assertEqual(selected_total, 0.0)
        self.assertEqual(maximum, 1.0)

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
