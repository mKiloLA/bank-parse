import math
import os
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional, Sequence

from .Utilities import get_appdata_path
from .data.DataTypes import Category, CategoryValue, category_from_name, category_name


@dataclass(frozen=True)
class CategoryRule:
    pattern: str
    title: str
    category: CategoryValue
    match_type: str = "contains"


@dataclass(frozen=True)
class CategoryTarget:
    name: str
    amount: float
    is_builtin: bool
    is_system: bool


class CategoryRuleStore:
    DUPLICATE_POLICIES = ("warn_review", "auto_merge", "ignore")
    DEFAULT_DUPLICATE_POLICY = "warn_review"

    def __init__(
        self,
        database_path: Optional[str] = None,
    ) -> None:
        self.database_path = database_path or get_appdata_path("statement_parser.db")
        self._memory_connection: Optional[sqlite3.Connection] = None
        if self.database_path != ":memory:":
            os.makedirs(os.path.dirname(os.path.abspath(self.database_path)), exist_ok=True)
        else:
            self._memory_connection = sqlite3.connect(":memory:")
            self._memory_connection.row_factory = sqlite3.Row

        self._initialize_schema()
        self._seed_categories()
        self._seed_category_targets()
        self._seed_app_settings()

    @staticmethod
    def normalize_pattern(pattern: str) -> str:
        return " ".join(pattern.strip().casefold().split())

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        connection = self._memory_connection
        owns_connection = connection is None
        if connection is None:
            connection = sqlite3.connect(self.database_path)
            connection.row_factory = sqlite3.Row
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            if owns_connection:
                connection.close()

    def _initialize_schema(self) -> None:
        with self._connection() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS category_rules (
                    pattern TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    category TEXT NOT NULL,
                    match_type TEXT NOT NULL
                        CHECK (match_type IN ('contains', 'exact'))
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS categories (
                    name TEXT PRIMARY KEY COLLATE NOCASE,
                    is_builtin INTEGER NOT NULL,
                    is_system INTEGER NOT NULL,
                    is_active INTEGER NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS category_targets (
                    category TEXT PRIMARY KEY,
                    amount REAL NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )

    @classmethod
    def _normalize_rule(
        cls, pattern: str, title: str, category: CategoryValue, match_type: str
    ) -> CategoryRule:
        normalized_pattern = cls.normalize_pattern(pattern)
        if not normalized_pattern:
            raise ValueError("A match pattern is required.")
        category_label = category_name(category).strip()
        if not category_label or category_label == Category.NONE.value:
            raise ValueError("Choose a category or Ignore for this rule.")
        if match_type not in ("contains", "exact"):
            raise ValueError("Match type must be 'contains' or 'exact'.")
        return CategoryRule(
            normalized_pattern,
            title.strip() or normalized_pattern,
            category_from_name(category_label),
            match_type,
        )

    @staticmethod
    def _upsert_rule(connection: sqlite3.Connection, rule: CategoryRule) -> None:
        connection.execute(
            """
            INSERT INTO category_rules(pattern, title, category, match_type)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(pattern) DO UPDATE SET
                title = excluded.title,
                category = excluded.category,
                match_type = excluded.match_type
            """,
            (
                rule.pattern,
                rule.title,
                category_name(rule.category),
                rule.match_type,
            ),
        )

    def _seed_categories(self) -> None:
        with self._connection() as connection:
            for category in Category:
                if category == Category.NONE:
                    continue
                connection.execute(
                    """
                    INSERT OR IGNORE INTO categories(
                        name, is_builtin, is_system, is_active
                    ) VALUES (?, 1, ?, 1)
                    """,
                    (
                        category.value,
                        int(category in (Category.INCOME, Category.IGNORE)),
                    ),
                )

    def _seed_category_targets(self) -> None:
        with self._connection() as connection:
            connection.execute(
                """
                INSERT OR IGNORE INTO category_targets(category, amount)
                SELECT name, 0 FROM categories
                WHERE is_active = 1 AND name != ?
                """,
                (Category.IGNORE.value,),
            )

    def _seed_app_settings(self) -> None:
        with self._connection() as connection:
            connection.execute(
                """
                INSERT OR IGNORE INTO app_settings(key, value)
                VALUES ('duplicate_policy', ?)
                """,
                (self.DEFAULT_DUPLICATE_POLICY,),
            )

    def get_duplicate_policy(self) -> str:
        with self._connection() as connection:
            row = connection.execute(
                "SELECT value FROM app_settings WHERE key = 'duplicate_policy'"
            ).fetchone()
        if row is None or row["value"] not in self.DUPLICATE_POLICIES:
            return self.DEFAULT_DUPLICATE_POLICY
        return row["value"]

    def set_duplicate_policy(self, policy: str) -> None:
        if policy not in self.DUPLICATE_POLICIES:
            raise ValueError("Choose a valid duplicate transaction policy.")
        with self._connection() as connection:
            connection.execute(
                """
                INSERT INTO app_settings(key, value)
                VALUES ('duplicate_policy', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (policy,),
            )

    @staticmethod
    def _active_category(
        connection: sqlite3.Connection, name: str
    ) -> Optional[sqlite3.Row]:
        return connection.execute(
            """
            SELECT name, is_builtin, is_system, is_active
            FROM categories
            WHERE name = ? COLLATE NOCASE AND is_active = 1
            """,
            (name,),
        ).fetchone()

    def list_rules(self) -> list[CategoryRule]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT pattern, title, category, match_type
                FROM category_rules
                ORDER BY pattern COLLATE NOCASE
                """
            ).fetchall()
        return [
            CategoryRule(
                row["pattern"],
                row["title"],
                category_from_name(row["category"]),
                row["match_type"],
            )
            for row in rows
        ]

    def list_categories(self) -> list[str]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT name FROM categories
                WHERE is_active = 1
                ORDER BY name COLLATE NOCASE
                """
            ).fetchall()
        return [row["name"] for row in rows]

    def list_category_targets(self) -> list[CategoryTarget]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT categories.name, category_targets.amount,
                       categories.is_builtin, categories.is_system
                FROM categories
                JOIN category_targets
                    ON category_targets.category = categories.name
                WHERE categories.is_active = 1
                ORDER BY categories.name COLLATE NOCASE
                """
            ).fetchall()
        return [
            CategoryTarget(
                row["name"],
                row["amount"],
                bool(row["is_builtin"]),
                bool(row["is_system"]),
            )
            for row in rows
        ]

    def save_rule(
        self,
        pattern: str,
        title: str,
        category: CategoryValue,
        match_type: str = "contains",
        original_pattern: Optional[str] = None,
    ) -> None:
        rule = self._normalize_rule(pattern, title, category, match_type)
        with self._connection() as connection:
            category_row = self._active_category(
                connection, category_name(rule.category)
            )
            if category_row is None:
                raise ValueError("Choose an existing category.")
            rule = CategoryRule(
                rule.pattern,
                rule.title,
                category_from_name(category_row["name"]),
                rule.match_type,
            )
            if original_pattern:
                old_pattern = self.normalize_pattern(original_pattern)
                if old_pattern != rule.pattern:
                    connection.execute(
                        "DELETE FROM category_rules WHERE pattern = ?", (old_pattern,)
                    )
            self._upsert_rule(connection, rule)

    def save_rules(self, rules: Sequence[CategoryRule]) -> None:
        normalized_rules = [
            self._normalize_rule(
                rule.pattern, rule.title, rule.category, rule.match_type
            )
            for rule in rules
        ]
        with self._connection() as connection:
            for rule in normalized_rules:
                category_row = self._active_category(
                    connection, category_name(rule.category)
                )
                if category_row is None:
                    raise ValueError("Choose an existing category.")
                rule = CategoryRule(
                    rule.pattern,
                    rule.title,
                    category_from_name(category_row["name"]),
                    rule.match_type,
                )
                self._upsert_rule(connection, rule)

    @staticmethod
    def _normalize_category_name(name: str) -> str:
        if not isinstance(name, str):
            raise ValueError("A category name is required.")
        normalized = name.strip()
        if not normalized or normalized.casefold() == Category.NONE.value.casefold():
            raise ValueError("A category name is required.")
        return normalized

    def save_category(
        self,
        name: str,
        amount: float,
        original_name: Optional[str] = None,
    ) -> None:
        normalized_name = self._normalize_category_name(name)
        if not math.isfinite(amount):
            raise ValueError("A category target must be a finite number.")

        with self._connection() as connection:
            if original_name is None:
                existing = connection.execute(
                    """
                    SELECT name, is_builtin, is_system, is_active
                    FROM categories WHERE name = ? COLLATE NOCASE
                    """,
                    (normalized_name,),
                ).fetchone()
                if existing is not None:
                    if not existing["is_builtin"] or existing["is_active"]:
                        raise ValueError(
                            f"The category {existing['name']!r} already exists."
                        )
                    connection.execute(
                        "UPDATE categories SET is_active = 1 WHERE name = ?",
                        (existing["name"],),
                    )
                    normalized_name = existing["name"]
                else:
                    connection.execute(
                        """
                        INSERT INTO categories(
                            name, is_builtin, is_system, is_active
                        ) VALUES (?, 0, 0, 1)
                        """,
                        (normalized_name,),
                    )
                connection.execute(
                    """
                    INSERT INTO category_targets(category, amount)
                    VALUES (?, ?)
                    ON CONFLICT(category) DO UPDATE SET amount = excluded.amount
                    """,
                    (normalized_name, amount),
                )
                return

            existing = self._active_category(connection, original_name)
            if existing is None:
                raise ValueError("The category no longer exists.")
            original_name = existing["name"]
            if existing["is_system"] and normalized_name != original_name:
                raise ValueError("System categories cannot be renamed.")

            if normalized_name != original_name:
                duplicate = connection.execute(
                    "SELECT 1 FROM categories WHERE name = ? COLLATE NOCASE",
                    (normalized_name,),
                ).fetchone()
                if duplicate:
                    raise ValueError(
                        f"The category {normalized_name!r} already exists."
                    )
                if existing["is_builtin"]:
                    connection.execute(
                        "UPDATE categories SET is_active = 0 WHERE name = ?",
                        (original_name,),
                    )
                    connection.execute(
                        """
                        INSERT INTO categories(
                            name, is_builtin, is_system, is_active
                        ) VALUES (?, 0, 0, 1)
                        """,
                        (normalized_name,),
                    )
                else:
                    connection.execute(
                        "UPDATE categories SET name = ? WHERE name = ?",
                        (normalized_name, original_name),
                    )
                connection.execute(
                    "UPDATE category_targets SET category = ? WHERE category = ?",
                    (normalized_name, original_name),
                )
                connection.execute(
                    "UPDATE category_rules SET category = ? WHERE category = ?",
                    (normalized_name, original_name),
                )
            connection.execute(
                "UPDATE category_targets SET amount = ? WHERE category = ?",
                (amount, normalized_name),
            )

    def delete_category(self, name: str) -> None:
        with self._connection() as connection:
            category = self._active_category(connection, name)
            if category is None:
                raise ValueError("The category no longer exists.")
            if category["is_system"]:
                raise ValueError("This system category cannot be deleted.")
            connection.execute(
                "DELETE FROM category_rules WHERE category = ?",
                (category["name"],),
            )
            connection.execute(
                "DELETE FROM category_targets WHERE category = ?",
                (category["name"],),
            )
            if category["is_builtin"]:
                connection.execute(
                    "UPDATE categories SET is_active = 0 WHERE name = ?",
                    (category["name"],),
                )
            else:
                connection.execute(
                    "DELETE FROM categories WHERE name = ?", (category["name"],)
                )

    def delete_rule(self, pattern: str) -> None:
        normalized_pattern = self.normalize_pattern(pattern)
        with self._connection() as connection:
            connection.execute(
                "DELETE FROM category_rules WHERE pattern = ?",
                (normalized_pattern,),
            )

    def get_category_targets(self) -> dict[str, float]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT category_targets.category, category_targets.amount
                FROM category_targets
                JOIN categories
                    ON categories.name = category_targets.category
                WHERE categories.is_active = 1 AND categories.name != ?
                ORDER BY category_targets.rowid
                """,
                (Category.IGNORE.value,),
            ).fetchall()
        return {row["category"]: row["amount"] for row in rows}

    def set_category_target(self, category: str, amount: float) -> None:
        if not math.isfinite(amount):
            raise ValueError("A category target must be a finite number.")
        with self._connection() as connection:
            category_row = self._active_category(connection, category)
            if category_row is None:
                raise ValueError(f"Unknown category {category!r}.")
            if category_row["name"] == Category.IGNORE.value:
                raise ValueError("This category does not have a monthly target.")
            connection.execute(
                """
                INSERT INTO category_targets(category, amount)
                VALUES (?, ?)
                ON CONFLICT(category) DO UPDATE SET amount = excluded.amount
                """,
                (category_row["name"], amount),
            )

    def close(self) -> None:
        if self._memory_connection is not None:
            self._memory_connection.close()

    def export_database(self, destination_path: str) -> None:
        destination_path = os.path.abspath(destination_path)
        if (
            self.database_path != ":memory:"
            and os.path.realpath(destination_path)
            == os.path.realpath(self.database_path)
        ):
            raise ValueError("Choose a different file from the active settings database.")

        os.makedirs(os.path.dirname(destination_path), exist_ok=True)
        destination = sqlite3.connect(destination_path)
        try:
            with self._connection() as source:
                source.backup(destination)
        finally:
            destination.close()

    def import_settings(self, source_path: str) -> None:
        source_path = os.path.abspath(source_path)
        if not os.path.isfile(source_path):
            raise ValueError("The selected settings file does not exist.")
        if (
            self.database_path != ":memory:"
            and os.path.realpath(source_path) == os.path.realpath(self.database_path)
        ):
            raise ValueError("Choose a settings file other than the active database.")

        source_uri = f"{Path(source_path).as_uri()}?mode=ro"
        try:
            source = sqlite3.connect(source_uri, uri=True)
        except sqlite3.DatabaseError as error:
            raise ValueError("The selected file is not a valid settings database.") from error

        try:
            source.row_factory = sqlite3.Row
            integrity = source.execute("PRAGMA integrity_check").fetchone()
            if integrity is None or integrity[0] != "ok":
                raise ValueError("The selected settings database is damaged.")

            required_columns = {
                "category_rules": {"pattern", "title", "category", "match_type"},
                "category_targets": {"category", "amount"},
            }
            for table, required in required_columns.items():
                columns = {
                    row["name"]
                    for row in source.execute(f"PRAGMA table_info({table})")
                }
                if not required <= columns:
                    raise ValueError(
                        "The selected file does not contain compatible Statement Parser settings."
                    )
            setting_columns = {
                row["name"] for row in source.execute("PRAGMA table_info(app_settings)")
            }
            imported_duplicate_policy = self.DEFAULT_DUPLICATE_POLICY
            if setting_columns:
                if not {"key", "value"} <= setting_columns:
                    raise ValueError(
                        "The selected file contains unsupported application settings."
                    )
                for row in source.execute(
                    "SELECT value FROM app_settings WHERE key = 'duplicate_policy'"
                ):
                    if row["value"] not in self.DUPLICATE_POLICIES:
                        raise ValueError(
                            "The selected file contains an invalid duplicate policy."
                        )
                    imported_duplicate_policy = row["value"]

            category_columns = {
                row["name"] for row in source.execute("PRAGMA table_info(categories)")
            }
            if category_columns:
                required = {"name", "is_builtin", "is_system", "is_active"}
                if not required <= category_columns:
                    raise ValueError(
                        "The selected file contains an unsupported category list."
                    )
                category_rows = source.execute(
                    """
                    SELECT name, is_builtin, is_system, is_active
                    FROM categories
                    """
                ).fetchall()
                imported_categories = {}
                for row in category_rows:
                    name = row["name"]
                    if not isinstance(name, str) or not name.strip():
                        raise ValueError(
                            "A category in the selected database is invalid."
                        )
                    key = name.casefold()
                    if key in imported_categories:
                        raise ValueError(
                            "The selected database contains duplicate categories."
                        )
                    imported_categories[key] = (
                        name,
                        int(bool(row["is_builtin"])),
                        int(bool(row["is_system"])),
                        int(bool(row["is_active"])),
                    )
            else:
                imported_categories = {
                    category.value.casefold(): (
                        category.value,
                        1,
                        int(category in (Category.INCOME, Category.IGNORE)),
                        1,
                    )
                    for category in Category
                    if category != Category.NONE
                }

            for required_system_category in (Category.INCOME.value, Category.IGNORE.value):
                record = imported_categories.get(required_system_category.casefold())
                if record is None or not record[2] or not record[3]:
                    raise ValueError(
                        "The settings file is missing a required system category."
                    )

            imported_rules = []
            seen_patterns = set()
            for row in source.execute(
                "SELECT pattern, title, category, match_type FROM category_rules"
            ):
                if not isinstance(row["pattern"], str) or not isinstance(
                    row["title"], str
                ):
                    raise ValueError("A rule in the selected database is invalid.")
                category_record = imported_categories.get(
                    str(row["category"]).casefold()
                )
                if category_record is None or not category_record[3]:
                    raise ValueError(
                        "A rule in the selected database uses an unavailable category."
                    )
                try:
                    category = category_from_name(category_record[0])
                    rule = self._normalize_rule(
                        row["pattern"],
                        row["title"],
                        category,
                        row["match_type"],
                    )
                except (TypeError, ValueError, AttributeError) as error:
                    raise ValueError(
                        "A rule in the selected database is invalid."
                    ) from error
                if rule.pattern in seen_patterns:
                    raise ValueError(
                        "The selected database contains duplicate rule patterns."
                    )
                seen_patterns.add(rule.pattern)
                imported_rules.append(rule)

            imported_targets = {}
            for row in source.execute(
                "SELECT category, amount FROM category_targets"
            ):
                category_record = imported_categories.get(
                    str(row["category"]).casefold()
                )
                if (
                    category_record is None
                    or not category_record[3]
                    or category_record[0] == Category.IGNORE.value
                    or category_record[0] in imported_targets
                ):
                    raise ValueError(
                        "A category target in the selected database is invalid."
                    )
                try:
                    amount = float(row["amount"])
                except (TypeError, ValueError) as error:
                    raise ValueError(
                        "A category target in the selected database is invalid."
                    ) from error
                if not math.isfinite(amount):
                    raise ValueError(
                        "A category target in the selected database is invalid."
                    )
                imported_targets[category_record[0]] = amount
        except sqlite3.DatabaseError as error:
            raise ValueError("The selected file is not a valid settings database.") from error
        finally:
            source.close()

        targets = {
            name: imported_targets.get(name, 0.0)
            for name, _is_builtin, _is_system, is_active in imported_categories.values()
            if is_active and name != Category.IGNORE.value
        }
        with self._connection() as connection:
            connection.execute("DELETE FROM category_rules")
            connection.execute("DELETE FROM category_targets")
            connection.execute("DELETE FROM categories")
            connection.execute("DELETE FROM app_settings")
            connection.executemany(
                """
                INSERT INTO categories(name, is_builtin, is_system, is_active)
                VALUES (?, ?, ?, ?)
                """,
                imported_categories.values(),
            )
            connection.executemany(
                """
                INSERT INTO category_rules(pattern, title, category, match_type)
                VALUES (?, ?, ?, ?)
                """,
                [
                    (
                        rule.pattern,
                        rule.title,
                        category_name(rule.category),
                        rule.match_type,
                    )
                    for rule in imported_rules
                ],
            )
            connection.executemany(
                "INSERT INTO category_targets(category, amount) VALUES (?, ?)",
                targets.items(),
            )
            connection.execute(
                "INSERT INTO app_settings(key, value) VALUES ('duplicate_policy', ?)",
                (imported_duplicate_policy,),
            )
