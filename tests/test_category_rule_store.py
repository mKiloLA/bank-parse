import os
import sqlite3
import tempfile
import unittest

from statement_parser.source.CategoryRuleStore import CategoryRuleStore
from statement_parser.source.data.DataTypes import Category


class CategoryRuleStoreTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.database_path = os.path.join(self.temporary_directory.name, "rules.db")

    def test_rules_and_targets_persist_between_store_instances(self):
        store = CategoryRuleStore(database_path=self.database_path)
        store.save_rule("coffee shop", "Coffee", Category.DINING)
        store.set_category_target(Category.DINING.value, 450.0)

        reopened_store = CategoryRuleStore(database_path=self.database_path)

        self.assertEqual(
            reopened_store.list_rules()[0].category, Category.DINING
        )
        self.assertEqual(
            reopened_store.get_category_targets()[Category.DINING.value], 450.0
        )

    def test_memory_database_keeps_data_between_store_operations(self):
        store = CategoryRuleStore(database_path=":memory:")
        self.addCleanup(store.close)

        store.save_rule("coffee", "Coffee", Category.DINING, "exact")

        self.assertEqual(store.list_rules()[0].match_type, "exact")
        store.set_category_target(Category.DINING.value, 125.0)
        self.assertEqual(store.get_category_targets()[Category.DINING.value], 125.0)

    def test_duplicate_policy_persists_and_rejects_unknown_values(self):
        store = CategoryRuleStore(database_path=self.database_path)
        self.assertEqual(store.get_duplicate_policy(), "warn_review")

        store.set_duplicate_policy("auto_merge")
        reopened_store = CategoryRuleStore(database_path=self.database_path)

        self.assertEqual(reopened_store.get_duplicate_policy(), "auto_merge")
        with self.assertRaisesRegex(ValueError, "valid duplicate"):
            reopened_store.set_duplicate_policy("remove_everything")

    def test_custom_categories_can_be_renamed_and_deleted_with_their_rules(self):
        store = CategoryRuleStore(database_path=self.database_path)
        store.save_category("Pet Care", 75.0)
        store.save_rule("pet store", "Pet Store", "Pet Care")

        store.save_category(
            "Animal Care", 90.0, original_name="Pet Care"
        )

        self.assertIn("Animal Care", store.get_category_targets())
        self.assertNotIn("Pet Care", store.get_category_targets())
        self.assertEqual(store.list_rules()[0].category, "Animal Care")

        store.delete_category("Animal Care")

        self.assertNotIn("Animal Care", store.list_categories())
        self.assertFalse(store.list_rules())

    def test_deleted_builtin_category_stays_deleted_across_restarts(self):
        store = CategoryRuleStore(database_path=self.database_path)
        store.delete_category(Category.DINING.value)

        reopened_store = CategoryRuleStore(database_path=self.database_path)

        self.assertNotIn(Category.DINING.value, reopened_store.list_categories())
        self.assertNotIn(Category.DINING.value, reopened_store.get_category_targets())
        with self.assertRaisesRegex(ValueError, "system category"):
            reopened_store.delete_category(Category.INCOME.value)

    def test_renaming_builtin_category_preserves_target_and_rules(self):
        store = CategoryRuleStore(database_path=self.database_path)
        store.save_category("Housing", 800.0, original_name=Category.MORTGAGE.value)
        store.save_rule("mortgage company", "Mortgage", "Housing")

        reopened_store = CategoryRuleStore(database_path=self.database_path)

        self.assertIn("Housing", reopened_store.get_category_targets())
        self.assertNotIn(Category.MORTGAGE.value, reopened_store.list_categories())
        self.assertEqual(reopened_store.list_rules()[0].category, "Housing")

    def test_exported_settings_replace_rules_and_targets_on_import(self):
        exporting_store = CategoryRuleStore(database_path=self.database_path)
        exporting_store.set_duplicate_policy("auto_merge")
        exporting_store.save_category("Pet Care", 450.0)
        exporting_store.save_rule("pet store", "Pet Store", "Pet Care", "exact")
        export_path = os.path.join(
            self.temporary_directory.name, "statement-parser-settings.db"
        )
        exporting_store.export_database(export_path)

        importing_store = CategoryRuleStore(
            database_path=os.path.join(self.temporary_directory.name, "current.db")
        )
        importing_store.save_rule("old merchant", "Old", Category.OTHER)
        importing_store.import_settings(export_path)

        self.assertEqual(
            [(rule.pattern, rule.category) for rule in importing_store.list_rules()],
            [("pet store", "Pet Care")],
        )
        self.assertEqual(
            importing_store.get_category_targets()["Pet Care"], 450.0
        )
        self.assertEqual(
            importing_store.get_category_targets()[Category.OTHER.value], 0.0
        )
        self.assertEqual(importing_store.get_duplicate_policy(), "auto_merge")

    def test_invalid_settings_file_does_not_change_current_database(self):
        store = CategoryRuleStore(database_path=self.database_path)
        store.save_rule("existing", "Existing", Category.OTHER)
        invalid_path = os.path.join(self.temporary_directory.name, "invalid.db")
        with open(invalid_path, "w", encoding="utf-8") as file:
            file.write("not a SQLite database")

        with self.assertRaisesRegex(ValueError, "valid settings database"):
            store.import_settings(invalid_path)

        self.assertEqual(
            [rule.pattern for rule in store.list_rules()], ["existing"]
        )

    def test_imports_settings_exported_before_category_registry_existed(self):
        old_export_path = os.path.join(
            self.temporary_directory.name, "old-settings.db"
        )
        with sqlite3.connect(old_export_path) as connection:
            connection.execute(
                """
                CREATE TABLE category_rules (
                    pattern TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    category TEXT NOT NULL,
                    match_type TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE category_targets (
                    category TEXT PRIMARY KEY,
                    amount REAL NOT NULL
                )
                """
            )
            connection.execute(
                "INSERT INTO category_rules VALUES (?, ?, ?, ?)",
                ("coffee", "Coffee", "Dining", "contains"),
            )
            connection.execute(
                "INSERT INTO category_targets VALUES (?, ?)", ("Dining", 300.0)
            )

        store = CategoryRuleStore(database_path=self.database_path)
        store.import_settings(old_export_path)

        self.assertIn(Category.DINING.value, store.list_categories())
        self.assertEqual(store.list_rules()[0].category, Category.DINING)
        self.assertEqual(
            store.get_category_targets()[Category.DINING.value], 300.0
        )
        self.assertEqual(store.get_duplicate_policy(), "warn_review")


if __name__ == "__main__":
    unittest.main()
