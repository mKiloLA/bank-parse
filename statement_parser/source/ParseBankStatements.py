import csv
import datetime
import math
import os
import re
from typing import Callable, Dict, Iterable, Optional, Sequence

import pandas as pd

from .CategoryRuleStore import CategoryRule, CategoryRuleStore
from .data.DataTypes import Bank, Category, CategoryValue
from .data.Statement import Statement
from .data.Transaction import Transaction


class ParseBankStatements:
    def __init__(
        self,
        logger: Callable[[str], None],
        database_path: Optional[str] = None,
    ) -> None:
        self.logger = logger
        self._rule_store = CategoryRuleStore(database_path=database_path)
        self._category_rules = self._rule_store.list_rules()

    @staticmethod
    def normalize_description(value: str) -> str:
        return " ".join(value.strip().casefold().split())

    @staticmethod
    def _normalize_header(value: str) -> str:
        return re.sub(r"[^a-z0-9]", "", value.casefold())

    def find_bank_statements(
        self, directory: str = "bank_statements"
    ) -> Dict[Bank, str]:
        """Find conventionally named statement files for older callers."""
        bank_files: Dict[Bank, str] = {}
        for bank in Bank:
            file_path = os.path.join(directory, f"{bank.value}.csv")
            if os.path.isfile(file_path):
                bank_files[bank] = file_path
                self.logger(f"Found statement for {bank.value}: {file_path}")
        return bank_files

    @staticmethod
    def _detect_bank(headers: Iterable[str]) -> Bank:
        normalized = {
            ParseBankStatements._normalize_header(header) for header in headers
        }
        if "datetime" in normalized and "amounttotal" in normalized:
            return Bank.VENMO
        if "transactiondate" in normalized and {"debit", "credit"} <= normalized:
            return Bank.CAPITAL_ONE
        if "transdate" in normalized and "amount" in normalized:
            return Bank.DISCOVER
        if "status" in normalized and {"date", "description"} <= normalized:
            return Bank.CITI
        if (
            "no" in normalized
            and {"date", "description", "debit", "credit"} <= normalized
        ):
            return Bank.COMMERCE
        return Bank.NONE

    @classmethod
    def _has_statement_headers(cls, headers: Iterable[str]) -> bool:
        normalized = {cls._normalize_header(header) for header in headers}
        if cls._detect_bank(headers) != Bank.NONE:
            return True

        has_date = bool(
            normalized
            & {"date", "transactiondate", "transdate", "datetime", "posteddate"}
        )
        has_description = bool(
            normalized & {"description", "merchant", "payee", "name"}
        )
        has_amount = bool(normalized & {"amount", "transactionamount"}) or bool(
            normalized & {"debit", "credit"}
        )
        return has_date and has_description and has_amount

    @classmethod
    def _find_header_row(cls, file_path: str) -> tuple[int, str]:
        with open(file_path, encoding="utf-8-sig", newline="") as file:
            sample = file.read(8192)
            file.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel

            for row_number, row in enumerate(csv.reader(file, dialect)):
                if cls._has_statement_headers(row):
                    return row_number, dialect.delimiter
                if row_number >= 99:
                    break

        raise ValueError(
            f"Could not locate a statement header in {os.path.basename(file_path)}. "
            "Expected bank columns or standard Date, Description, and Amount columns."
        )

    @staticmethod
    def _column(
        header_map: Dict[str, str],
        choices: Iterable[str],
        label: str,
        file_path: str,
    ) -> str:
        for choice in choices:
            original = header_map.get(choice)
            if original is not None:
                return original
        raise ValueError(
            f"Could not find a {label} column in {os.path.basename(file_path)}. "
            "Expected standard Date, Description, and Amount columns or a supported bank format."
        )

    @staticmethod
    def _parse_date(value: str, file_path: str, row_number: int) -> datetime.date:
        text = str(value).strip()
        if not text:
            raise ValueError(
                f"{os.path.basename(file_path)} row {row_number}: transaction date is empty."
            )
        try:
            return datetime.date.fromisoformat(text[:10])
        except ValueError:
            pass

        for date_format in ("%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(text, date_format).date()
            except ValueError:
                continue
        raise ValueError(
            f"{os.path.basename(file_path)} row {row_number}: "
            f"unrecognized transaction date {text!r}."
        )

    @staticmethod
    def _parse_amount(value: str, file_path: str, row_number: int) -> float:
        text = str(value).strip().replace("\u2212", "-")
        if not text:
            raise ValueError(
                f"{os.path.basename(file_path)} row {row_number}: transaction amount is empty."
            )

        negative_parentheses = text.startswith("(") and text.endswith(")")
        if negative_parentheses:
            text = text[1:-1].strip()
        text = text.replace("$", "").replace(",", "").replace(" ", "")
        try:
            amount = float(text)
        except ValueError as error:
            raise ValueError(
                f"{os.path.basename(file_path)} row {row_number}: "
                f"unrecognized transaction amount {value!r}."
            ) from error
        if not math.isfinite(amount):
            raise ValueError(
                f"{os.path.basename(file_path)} row {row_number}: "
                f"unrecognized transaction amount {value!r}."
            )
        return -abs(amount) if negative_parentheses else amount

    @staticmethod
    def _empty(value: str) -> bool:
        return not str(value).strip()

    def _find_category_rule(
        self, description: str
    ) -> Optional[tuple[str, CategoryValue]]:
        matches = [
            rule
            for rule in self._category_rules
            if self.category_rule_matches(rule.pattern, description, rule.match_type)
        ]
        match = max(
            matches,
            key=lambda rule: (rule.match_type == "exact", len(rule.pattern)),
            default=None,
        )
        return (match.title, match.category) if match else None

    @classmethod
    def category_rule_matches(
        cls, pattern: str, description: str, match_type: str = "contains"
    ) -> bool:
        normalized_pattern = cls.normalize_description(pattern)
        normalized_description = cls.normalize_description(description)
        if match_type == "exact":
            return normalized_pattern == normalized_description
        if match_type == "contains":
            return normalized_pattern in normalized_description
        raise ValueError("Match type must be 'contains' or 'exact'.")

    def apply_category_rules(self, statements: Sequence[Statement]) -> None:
        for statement in statements:
            for transaction in statement.transactions:
                rule = self._find_category_rule(transaction.description)
                if rule:
                    transaction.title, transaction.category = rule

    def save_category_rule(
        self,
        pattern: str,
        title: str,
        category: CategoryValue,
        match_type: str = "contains",
        original_pattern: Optional[str] = None,
    ) -> None:
        self._rule_store.save_rule(
            pattern, title, category, match_type, original_pattern
        )
        self._category_rules = self._rule_store.list_rules()

    def save_category_rules(
        self,
        rules: Dict[str, tuple[str, CategoryValue]],
        match_type: str = "contains",
    ) -> None:
        self._rule_store.save_rules(
            [
                CategoryRule(
                    pattern=pattern,
                    title=title,
                    category=category,
                    match_type=match_type,
                )
                for pattern, (title, category) in rules.items()
            ]
        )
        self._category_rules = self._rule_store.list_rules()

    @property
    def category_rules(self) -> list[CategoryRule]:
        return list(self._category_rules)

    def delete_category_rule(self, pattern: str) -> None:
        self._rule_store.delete_rule(pattern)
        self._category_rules = self._rule_store.list_rules()

    @property
    def categories(self) -> list[str]:
        return self._rule_store.list_categories()

    def save_category(
        self, name: str, amount: float, original_name: Optional[str] = None
    ) -> None:
        self._rule_store.save_category(name, amount, original_name)
        self._category_rules = self._rule_store.list_rules()

    def delete_category(self, name: str) -> None:
        self._rule_store.delete_category(name)
        self._category_rules = self._rule_store.list_rules()

    def export_settings(self, destination_path: str) -> None:
        self._rule_store.export_database(destination_path)

    def import_settings(self, source_path: str) -> None:
        self._rule_store.import_settings(source_path)
        self._category_rules = self._rule_store.list_rules()

    @property
    def duplicate_policy(self) -> str:
        return self._rule_store.get_duplicate_policy()

    def set_duplicate_policy(self, policy: str) -> None:
        self._rule_store.set_duplicate_policy(policy)

    def statement_from_file(self, bank: Optional[Bank], file_path: str) -> Statement:
        """Parse a bank CSV using its headers, or standard Date/Description/Amount columns."""
        self.logger(f"Reading statement: {file_path}")
        header_row, delimiter = self._find_header_row(file_path)
        try:
            frame = pd.read_csv(
                file_path,
                sep=delimiter,
                engine="python",
                encoding="utf-8-sig",
                skiprows=header_row,
                header=0,
                dtype=str,
                keep_default_na=False,
            )
        except (
            pd.errors.ParserError,
            pd.errors.EmptyDataError,
            UnicodeDecodeError,
        ) as error:
            raise ValueError(
                f"Could not read {os.path.basename(file_path)}: {error}"
            ) from error

        if frame.empty and len(frame.columns) == 0:
            raise ValueError(
                f"{os.path.basename(file_path)} does not contain a CSV header."
            )

        original_headers = [str(column).strip() for column in frame.columns]
        header_map = {
            self._normalize_header(header): header for header in original_headers
        }
        detected_bank = self._detect_bank(original_headers)
        bank = bank if bank not in (None, Bank.NONE) else detected_bank

        if bank == Bank.VENMO:
            date_column = self._column(header_map, ("datetime",), "date", file_path)
            amount_column = self._column(
                header_map, ("amounttotal",), "amount", file_path
            )
            description_column = header_map.get("note")
            debit_credit = False
            invert_amount = False
        elif bank == Bank.DISCOVER:
            date_column = self._column(header_map, ("transdate",), "date", file_path)
            amount_column = self._column(header_map, ("amount",), "amount", file_path)
            description_column = self._column(
                header_map, ("description",), "description", file_path
            )
            debit_credit = False
            invert_amount = True
        elif bank == Bank.CAPITAL_ONE:
            date_column = self._column(
                header_map, ("transactiondate",), "date", file_path
            )
            amount_column = ""
            description_column = self._column(
                header_map, ("description",), "description", file_path
            )
            debit_credit = True
            invert_amount = False
        elif bank in (Bank.COMMERCE, Bank.CITI):
            date_column = self._column(header_map, ("date",), "date", file_path)
            amount_column = ""
            description_column = self._column(
                header_map, ("description",), "description", file_path
            )
            debit_credit = True
            invert_amount = False
        else:
            date_column = self._column(
                header_map,
                ("date", "transactiondate", "transdate", "datetime", "posteddate"),
                "date",
                file_path,
            )
            description_column = self._column(
                header_map,
                ("description", "merchant", "payee", "name"),
                "description",
                file_path,
            )
            amount_column = header_map.get("amount") or header_map.get(
                "transactionamount"
            )
            debit_credit = not amount_column and (
                "debit" in header_map or "credit" in header_map
            )
            if not amount_column and not debit_credit:
                raise ValueError(
                    f"Could not find an amount column in {os.path.basename(file_path)}. "
                    "Expected an Amount column or Debit/Credit columns."
                )
            invert_amount = False
            if bank in (None, Bank.NONE):
                bank = detected_bank

        transactions = []
        for index, row in frame.iterrows():
            values = {
                str(key).strip(): str(value).strip() for key, value in row.items()
            }
            if not any(values.values()):
                continue
            row_number = int(index) + header_row + 2
            if (
                bank == Bank.VENMO
                and self._empty(values[date_column])
                and any(
                    column in header_map
                    and not self._empty(values.get(header_map[column], ""))
                    for column in (
                        "beginningbalance",
                        "endingbalance",
                        "statementperiodvenmofees",
                        "yeartodatevenmofees",
                        "disclaimer",
                    )
                )
            ):
                continue
            date = self._parse_date(values[date_column], file_path, row_number)

            if debit_credit:
                debit = values.get(header_map.get("debit", ""), "")
                credit = values.get(header_map.get("credit", ""), "")
                if self._empty(debit) and self._empty(credit):
                    raise ValueError(
                        f"{os.path.basename(file_path)} row {row_number}: "
                        "both Debit and Credit amounts are empty."
                    )
                amount = 0.0
                if not self._empty(debit):
                    amount -= abs(self._parse_amount(debit, file_path, row_number))
                if not self._empty(credit):
                    amount += abs(self._parse_amount(credit, file_path, row_number))
            else:
                amount = self._parse_amount(
                    values[amount_column], file_path, row_number
                )
                if invert_amount:
                    amount *= -1

            description = values.get(description_column or "", "").strip()
            if bank == Bank.VENMO and not description:
                from_person = values.get(header_map.get("from", ""), "").strip()
                to_person = values.get(header_map.get("to", ""), "").strip()
                description = f"Venmo: {from_person} -> {to_person}".strip()
            if not description:
                raise ValueError(
                    f"{os.path.basename(file_path)} row {row_number}: "
                    "transaction description is empty."
                )

            rule = self._find_category_rule(description)
            title, category = rule if rule else (description, Category.NONE)
            transactions.append(
                Transaction(
                    date=date,
                    amount=amount,
                    description=description,
                    bank=bank or Bank.NONE,
                    category=category,
                    title=title,
                )
            )

        self.logger(
            f"Previewed {len(transactions)} transactions from "
            f"{os.path.basename(file_path)} ({bank.value})."
        )
        return Statement(bank=bank or Bank.NONE, transactions=transactions)
