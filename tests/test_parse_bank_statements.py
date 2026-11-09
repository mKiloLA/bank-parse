import os
import tempfile
import unittest

from statement_parser.source.ParseBankStatements import ParseBankStatements
from statement_parser.source.CategoryTargetHandler import CategoryTargetHandler
from statement_parser.source.data.DataTypes import Bank, Category


class ParseBankStatementsTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.database_path = os.path.join(
            self.temporary_directory.name, "category-rules.db"
        )
        self.messages = []
        self.parser = ParseBankStatements(
            self.messages.append, database_path=self.database_path
        )

    def write_csv(self, filename, content):
        path = os.path.join(self.temporary_directory.name, filename)
        with open(path, "w", encoding="utf-8", newline="") as file:
            file.write(content)
        return path

    def test_detects_discover_from_headers_and_parses_file_without_bank_name(self):
        path = self.write_csv(
            "download.csv",
            "Trans. Date,Post Date,Description,Amount,Category\n"
            "01/15/2025,01/16/2025,Local Market,42.15,Supermarkets\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.DISCOVER)
        self.assertEqual(statement.transactions[0].amount, -42.15)
        self.assertEqual(statement.transactions[0].description, "Local Market")
        self.assertEqual(statement.transactions[0].category, Category.NONE)

    def test_parses_generic_headers_and_parenthesized_amounts(self):
        path = self.write_csv(
            "my-export.csv",
            "Transaction Date,Merchant,Amount\n" "2025-02-01,Corner Store,(12.50)\n",
        )

        statement = self.parser.statement_from_file(None, path)

        transaction = statement.transactions[0]
        self.assertEqual(statement.bank, Bank.NONE)
        self.assertEqual(transaction.date.isoformat(), "2025-02-01")
        self.assertEqual(transaction.amount, -12.5)
        self.assertEqual(transaction.description, "Corner Store")

    def test_detects_commerce_and_handles_split_debit_credit_columns(self):
        path = self.write_csv(
            "statement.csv",
            "Date,No.,Description,Debit,Credit\n"
            "01/01/2025,1,Hardware Shop,19.99,\n"
            "01/02/2025,2,Paycheck,,100.00\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.COMMERCE)
        self.assertEqual(
            [transaction.amount for transaction in statement.transactions],
            [-19.99, 100.0],
        )

    def test_detects_capital_one_from_transaction_date_and_debit_credit(self):
        path = self.write_csv(
            "download.csv",
            "Transaction Date,Posted Date,Card No.,Description,Category,Debit,Credit\n"
            "2025-01-03,2025-01-04,1234,Cafe,Food,8.50,\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.CAPITAL_ONE)
        self.assertEqual(statement.transactions[0].amount, -8.5)

    def test_detects_citi_from_status_and_date_headers(self):
        path = self.write_csv(
            "download.csv",
            "Status,Date,Description,Debit,Credit\n"
            "Cleared,1/03/2025,Utility Company,80.00,\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.CITI)
        self.assertEqual(statement.transactions[0].amount, -80.0)

    def test_detects_venmo_and_uses_people_when_note_is_empty(self):
        path = self.write_csv(
            "download.csv",
            "ID,Datetime,Type,Status,Note,From,To,Amount (total)\n"
            "1,2025-01-04T10:00:00,payment,Complete,,Alex,Sam,-14.00\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.VENMO)
        self.assertEqual(statement.transactions[0].amount, -14.0)
        self.assertEqual(statement.transactions[0].description, "Venmo: Alex -> Sam")

    def test_detects_venmo_statement_after_account_header_preamble(self):
        # Venmo statements include balance summary rows alongside transactions.
        path = self.write_csv(
            "VenmoStatement_September_2026.csv",
            "Account Statement\n\n"
            "ID,Datetime,Type,Status,Note,From,To,Amount (total),Beginning Balance,Ending Balance\n"
            ",,,,,,,,100.00,\n"
            "1,2026-09-10 10:15:00,Payment,Complete,Local Market,Alex,Sam,-$12.50,,\n"
            ",,,,,,,,,87.50\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.bank, Bank.VENMO)
        self.assertEqual(len(statement.transactions), 1)
        self.assertEqual(statement.transactions[0].date.isoformat(), "2026-09-10")
        self.assertEqual(statement.transactions[0].amount, -12.5)
        self.assertEqual(statement.transactions[0].description, "Local Market")

    def test_uses_the_most_specific_existing_merchant_rule(self):
        self.parser.save_category_rule("coffee", "Coffee", Category.DINING)
        self.parser.save_category_rule(
            "coffee shop", "Coffee Shop", Category.SUPERMARKET
        )
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "2025-03-01,COFFEE SHOP #123,8.00\n",
        )

        statement = self.parser.statement_from_file(None, path)

        transaction = statement.transactions[0]
        self.assertEqual(transaction.category, Category.SUPERMARKET)
        self.assertEqual(transaction.title, "Coffee Shop")

    def test_bulk_rule_can_match_a_short_keyword_in_long_descriptions(self):
        self.parser.save_category_rules(
            {"google": ("Google", Category.UTILITIES)},
            match_type="contains",
        )
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n"
            "2025-03-01,ACH Debit SERVICES GOOGLE 0910000012807524,-80.00\n"
            "2025-03-02,GOOGLE FIBER PAYMENT,-60.00\n"
            "2025-03-03,Other Utility Company,-20.00\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(
            [transaction.category for transaction in statement.transactions],
            [Category.UTILITIES, Category.UTILITIES, Category.NONE],
        )
        self.assertEqual(
            [transaction.title for transaction in statement.transactions],
            ["Google", "Google", "Other Utility Company"],
        )

    def test_new_broad_rule_updates_matching_transactions_already_loaded(self):
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n"
            "2025-03-01,ACH Debit SERVICES GOOGLE 0910000012807524,-80.00\n"
            "2025-03-02,GOOGLE FIBER PAYMENT,-60.00\n"
            "2025-03-03,Other Utility Company,-20.00\n",
        )
        statement = self.parser.statement_from_file(None, path)
        self.parser.save_category_rules(
            {"google": ("Google", Category.UTILITIES)},
            match_type="contains",
        )

        self.parser.apply_category_rules([statement])

        self.assertEqual(
            [transaction.category for transaction in statement.transactions],
            [Category.UTILITIES, Category.UTILITIES, Category.NONE],
        )

    def test_new_rules_are_persisted_and_loaded_on_next_run(self):
        self.parser.save_category_rule(
            "  Local   Market ", "Groceries", Category.SUPERMARKET
        )
        reloaded_parser = ParseBankStatements(
            self.messages.append, database_path=self.database_path
        )
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "2025-03-01,Local Market #52,8.00\n",
        )

        transaction = reloaded_parser.statement_from_file(None, path).transactions[0]

        self.assertEqual(transaction.category, Category.SUPERMARKET)
        self.assertEqual(transaction.title, "Groceries")

    def test_custom_category_is_available_to_rules_and_imported_transactions(self):
        self.parser.save_category("Pet Care", 250.0)
        self.parser.save_category_rule("pet supply", "Pet Supplies", "Pet Care")
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "2025-03-01,Pet Supply Shop,-18.00\n",
        )

        transaction = self.parser.statement_from_file(None, path).transactions[0]

        self.assertIn("Pet Care", self.parser.categories)
        self.assertEqual(transaction.category, "Pet Care")
        self.assertTrue(transaction.is_purchase())

    def test_exact_rules_only_match_the_complete_normalized_description(self):
        self.parser.save_category_rule(
            "Local Market",
            "Local Market",
            Category.SUPERMARKET,
            match_type="exact",
        )
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n"
            "2025-03-01, LOCAL   MARKET ,-8.00\n"
            "2025-03-02,Local Market #52,-9.00\n",
        )

        statement = self.parser.statement_from_file(None, path)

        self.assertEqual(statement.transactions[0].category, Category.SUPERMARKET)
        self.assertEqual(statement.transactions[1].category, Category.NONE)

    def test_rule_can_be_edited_and_deleted(self):
        self.parser.save_category_rule(
            "coffee", "Coffee", Category.DINING, match_type="exact"
        )
        self.parser.save_category_rule(
            "cafeteria",
            "Cafe",
            Category.OTHER,
            match_type="contains",
            original_pattern="coffee",
        )
        self.assertEqual(
            [(rule.pattern, rule.match_type) for rule in self.parser.category_rules],
            [("cafeteria", "contains")],
        )

        self.parser.delete_category_rule("CAFETERIA")

        self.assertEqual(self.parser.category_rules, [])

    def test_ignore_rules_are_reused_and_not_added_to_monthly_targets(self):
        self.parser.save_category_rule(
            "internal transfer", "Internal transfer", Category.IGNORE
        )
        parser = ParseBankStatements(
            self.messages.append, database_path=self.database_path
        )
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "2025-03-01,Internal Transfer,200.00\n",
        )
        target_path = os.path.join(self.temporary_directory.name, "targets.db")
        targets = CategoryTargetHandler(target_path)

        transaction = parser.statement_from_file(None, path).transactions[0]

        self.assertEqual(transaction.category, Category.IGNORE)
        self.assertNotIn(Category.IGNORE.value, targets.targets)

    def test_invalid_date_is_reported_instead_of_replaced_with_today(self):
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "not-a-date,Corner Store,12.50\n",
        )

        with self.assertRaisesRegex(ValueError, "row 2.*unrecognized transaction date"):
            self.parser.statement_from_file(None, path)

    def test_non_finite_amount_is_rejected(self):
        path = self.write_csv(
            "generic.csv",
            "Date,Description,Amount\n" "2025-03-01,Corner Store,nan\n",
        )

        with self.assertRaisesRegex(
            ValueError, "row 2.*unrecognized transaction amount"
        ):
            self.parser.statement_from_file(None, path)


if __name__ == "__main__":
    unittest.main()
