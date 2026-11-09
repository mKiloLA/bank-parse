import os
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from typing import Optional, Sequence

from openpyxl import load_workbook

from .data.DataTypes import Bank, Category, category_from_name, category_name
from .data.Statement import Statement
from .data.Transaction import Transaction


@dataclass(frozen=True)
class DuplicateOccurrence:
    statement_index: int
    transaction_index: int
    source_path: str
    source_key: str
    transaction: Transaction


@dataclass(frozen=True)
class DuplicateGroup:
    occurrences: tuple[DuplicateOccurrence, ...]


@dataclass(frozen=True)
class CategoryHistory:
    category: str
    average_monthly_spend: float
    monthly_change: Optional[float]
    monthly_spending: tuple[float, ...]

    @property
    def first_spending_month_index(self) -> int:
        return _first_spending_month_index(self.monthly_spending)


@dataclass(frozen=True)
class SubscriptionCandidate:
    merchant: str
    amount: float
    months: tuple[date, ...]
    transaction_count: int
    first_date: date
    last_date: date


@dataclass(frozen=True)
class HistoryAnalysis:
    months: tuple[date, ...]
    monthly_income: tuple[float, ...]
    monthly_spending: tuple[float, ...]
    category_history: tuple[CategoryHistory, ...]
    subscriptions: tuple[SubscriptionCandidate, ...]

    @property
    def total_income(self) -> float:
        return sum(self.monthly_income)

    @property
    def total_spending(self) -> float:
        return sum(self.monthly_spending)

    @property
    def first_spending_month_index(self) -> int:
        return _first_spending_month_index(self.monthly_spending)

    @property
    def monthly_spending_change(self) -> Optional[float]:
        return _monthly_slope(self.monthly_spending[self.first_spending_month_index :])


def normalize_place(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value).casefold()
    without_marks = "".join(
        character for character in decomposed if not unicodedata.combining(character)
    )
    return " ".join(re.sub(r"[^\w]+", " ", without_marks).split())


def duplicate_key(transaction: Transaction) -> Optional[tuple[date, float, str]]:
    place = normalize_place(transaction.title or transaction.description or "")
    if not place:
        return None
    return (
        transaction.date,
        round(float(transaction.amount), 2),
        place,
    )


def _source_key(path: str, index: int) -> str:
    return os.path.normcase(os.path.realpath(path)) if path else f"source-{index}"


def find_duplicate_groups(
    statements: Sequence[Statement], source_files: Sequence[str]
) -> list[DuplicateGroup]:
    grouped = defaultdict(list)
    for statement_index, statement in enumerate(statements):
        source_path = (
            source_files[statement_index] if statement_index < len(source_files) else ""
        )
        source_key = _source_key(source_path, statement_index)
        for transaction_index, transaction in enumerate(statement.transactions):
            key = duplicate_key(transaction)
            if key is not None:
                grouped[key].append(
                    DuplicateOccurrence(
                        statement_index,
                        transaction_index,
                        source_path,
                        source_key,
                        transaction,
                    )
                )
    return [
        DuplicateGroup(tuple(occurrences))
        for occurrences in grouped.values()
        if len({occurrence.source_key for occurrence in occurrences}) > 1
    ]


def duplicate_occurrences_to_skip(
    groups: Sequence[DuplicateGroup],
) -> set[tuple[int, int]]:
    excluded = set()
    for group in groups:
        first_source = group.occurrences[0].source_key
        excluded.update(
            (occurrence.statement_index, occurrence.transaction_index)
            for occurrence in group.occurrences
            if occurrence.source_key != first_source
        )
    return excluded


def filter_duplicate_occurrences(
    statements: Sequence[Statement],
    excluded: set[tuple[int, int]],
) -> list[Statement]:
    return [
        Statement(
            statement.bank,
            [
                transaction
                for transaction_index, transaction in enumerate(statement.transactions)
                if (statement_index, transaction_index) not in excluded
            ],
        )
        for statement_index, statement in enumerate(statements)
    ]


def _parse_workbook_date(value, path: str, row_number: int) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for parser in (
            lambda item: date.fromisoformat(item[:10]),
            lambda item: datetime.strptime(item, "%m/%d/%Y").date(),
            lambda item: datetime.strptime(item, "%m/%d/%y").date(),
        ):
            try:
                return parser(text)
            except ValueError:
                continue
    raise ValueError(
        f"{os.path.basename(path)} row {row_number}: invalid transaction date."
    )


def _parse_workbook_amount(value, path: str, row_number: int) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace("$", "").replace(",", "")
        if text.startswith("(") and text.endswith(")"):
            text = f"-{text[1:-1]}"
        try:
            return float(text)
        except ValueError:
            pass
    raise ValueError(
        f"{os.path.basename(path)} row {row_number}: invalid transaction amount."
    )


def load_combined_statement(path: str) -> Statement:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Expenses & Income" not in workbook.sheetnames:
            raise ValueError(
                f"{os.path.basename(path)} is missing the Expenses & Income sheet."
            )
        sheet = workbook["Expenses & Income"]
        tables = []
        for row_number in range(1, min(sheet.max_row, 12) + 1):
            for start_column in (2, 8):
                headers = [
                    str(sheet.cell(row_number, start_column + offset).value or "")
                    .strip()
                    .casefold()
                    for offset in range(4)
                ]
                if headers == ["title", "date", "category", "amount"]:
                    tables.append((row_number, start_column))
        if not tables:
            raise ValueError(
                f"{os.path.basename(path)} does not contain a recognized transaction table."
            )

        transactions = []
        for header_row, start_column in tables:
            for row_number in range(header_row + 1, sheet.max_row + 1):
                title = sheet.cell(row_number, start_column).value
                date_value = sheet.cell(row_number, start_column + 1).value
                category_value = sheet.cell(row_number, start_column + 2).value
                amount_value = sheet.cell(row_number, start_column + 3).value
                if all(
                    value is None
                    for value in (title, date_value, category_value, amount_value)
                ):
                    continue
                transaction_date = _parse_workbook_date(date_value, path, row_number)
                amount = _parse_workbook_amount(amount_value, path, row_number)
                title_text = str(title or "").strip()
                category = category_from_name(str(category_value or "").strip())
                transactions.append(
                    Transaction(
                        transaction_date,
                        amount,
                        title_text,
                        Bank.NONE,
                        category,
                        title_text,
                    )
                )
        if not transactions:
            raise ValueError(f"{os.path.basename(path)} contains no transaction rows.")
        return Statement(Bank.NONE, transactions)
    finally:
        workbook.close()


def _monthly_slope(values: Sequence[float]) -> Optional[float]:
    if len(values) < 2:
        return None
    average_x = (len(values) - 1) / 2
    average_y = sum(values) / len(values)
    denominator = sum((index - average_x) ** 2 for index in range(len(values)))
    return (
        sum(
            (index - average_x) * (value - average_y)
            for index, value in enumerate(values)
        )
        / denominator
    )


def _first_spending_month_index(values: Sequence[float]) -> int:
    return next((index for index, value in enumerate(values) if value > 0), 0)


def analyze_history(statements: Sequence[Statement]) -> HistoryAnalysis:
    transactions = [
        transaction
        for statement in statements
        for transaction in statement.transactions
    ]
    if not transactions:
        return HistoryAnalysis((), (), (), (), ())

    first_date = min(transaction.date for transaction in transactions)
    last_date = max(transaction.date for transaction in transactions)
    first_month = date(first_date.year, first_date.month, 1)
    last_month = date(last_date.year, last_date.month, 1)
    month_count = (
        (last_month.year - first_month.year) * 12
        + last_month.month
        - first_month.month
        + 1
    )
    months = tuple(
        date(
            first_month.year + (first_month.month - 1 + offset) // 12,
            (first_month.month - 1 + offset) % 12 + 1,
            1,
        )
        for offset in range(month_count)
    )
    monthly_categories = defaultdict(lambda: defaultdict(float))
    monthly_incomes = defaultdict(float)
    subscription_groups = defaultdict(list)
    category_names = set()

    for transaction in transactions:
        month = date(transaction.date.year, transaction.date.month, 1)
        category = category_name(transaction.category)
        if category in (Category.NONE.value, Category.IGNORE.value):
            continue
        if category == Category.INCOME.value:
            monthly_incomes[month] += float(transaction.amount)
            continue
        monthly_categories[month][category] += float(transaction.amount)
        category_names.add(category)
        if transaction.amount < 0:
            place = transaction.title or transaction.description or ""
            normalized_place = normalize_place(place)
            if normalized_place:
                subscription_groups[
                    (normalized_place, round(abs(float(transaction.amount)), 2))
                ].append(transaction)

    category_history = []
    category_values = {}
    for category in category_names:
        values = tuple(
            abs(monthly_categories.get(month, {}).get(category, 0.0))
            for month in months
        )
        active_values = values[_first_spending_month_index(values) :]
        category_values[category] = values
        category_history.append(
            CategoryHistory(
                category,
                sum(active_values) / len(active_values),
                _monthly_slope(active_values),
                values,
            )
        )
    category_history.sort(
        key=lambda item: (-item.average_monthly_spend, item.category.casefold())
    )

    monthly_income = tuple(abs(monthly_incomes.get(month, 0.0)) for month in months)
    monthly_spending = tuple(
        sum(category_values[category][index] for category in category_values)
        for index in range(len(months))
    )
    subscriptions = []
    for (
        normalized_place,
        amount,
    ), matching_transactions in subscription_groups.items():
        months_seen = tuple(
            sorted(
                {
                    date(item.date.year, item.date.month, 1)
                    for item in matching_transactions
                }
            )
        )
        if len(months_seen) < 3:
            continue
        first_transaction = min(matching_transactions, key=lambda item: item.date)
        subscriptions.append(
            SubscriptionCandidate(
                first_transaction.title
                or first_transaction.description
                or normalized_place,
                amount,
                months_seen,
                len(matching_transactions),
                min(item.date for item in matching_transactions),
                max(item.date for item in matching_transactions),
            )
        )
    subscriptions.sort(
        key=lambda item: (-len(item.months), -item.amount, item.merchant.casefold())
    )
    return HistoryAnalysis(
        months,
        monthly_income,
        monthly_spending,
        tuple(category_history),
        tuple(subscriptions),
    )
