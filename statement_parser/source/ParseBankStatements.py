import csv
import datetime
import math
import os
import pandas as pd
from typing import Dict, Callable

from .data.DataTypes import Bank, Category
from .data.Statement import Statement
from .data.Transaction import Transaction
from .Utilities import get_appdata_path


# Load descriptionToCategory.csv into a lookup dict once per run
DescriptionToCategoryPath = get_appdata_path('descriptionToCategory.csv')
desc_cat_map: dict[str, tuple[str, Category]] = {}
if not os.path.isfile(DescriptionToCategoryPath):
    with open(DescriptionToCategoryPath, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=',', lineterminator='\n')
        writer.writerow(['description', 'title', 'category'])
        
if os.path.isfile(DescriptionToCategoryPath):
    with open(DescriptionToCategoryPath, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            desc = row.get('description', '').strip()
            title = row.get('title', '').strip()
            cat_str = row.get('category', '').strip()
            cat_enum = next((c for c in Category if c.value == cat_str), Category.NONE)
            desc_cat_map[desc] = (title, cat_enum)

class ParseBankStatements:
    def __init__(self, logger: Callable[[str], None], get_category_callback: Callable[[str, float, str], tuple[str, str, str]]) -> None:
        self.logger: Callable[[str], None] = logger
        self.get_category_callback = get_category_callback

    def find_bank_statements(self, directory: str = "bank_statements") -> Dict[Bank, str]:
        """
        Returns a dict mapping Bank enum to the file path of its statement CSV, if present.
        """
        bank_files: Dict[Bank, str] = {}
        for bank in Bank:
            bank_filename = f"{bank.value}.csv"
            file_path = os.path.join(directory, bank_filename)

            if os.path.isfile(file_path):
                bank_files[bank] = file_path
                self.logger(f"Found statement for {bank.value}: {file_path}")
        return bank_files


    def statement_from_file(self, bank: Bank, file_path: str) -> Statement:
        """
        Reads the given CSV file and returns a Statement object populated with Transaction objects for the given bank.
        Assumes columns: Date, Description, Amount (or Debit/Credit), Title, Category
        """
        self.logger(f"Parsing file for {bank.value}: {file_path}")
        df = pd.read_csv(file_path)
        transactions = []
        for _, row in df.iterrows():
            match bank:
                # Commerce.csv headers:
                #   Date, No., Description, Debit, Credit
                #   No explicit category column. All transactions default to Category.NONE.
                case Bank.COMMERCE:
                    try:
                        date = datetime.datetime.strptime(str(row.get("Date", "")), "%m/%d/%Y").date()
                    except Exception:
                        date = None

                    debit = row.get("Debit")
                    credit = row.get("Credit")
                    if debit and not math.isnan(debit):
                        amount = -float(debit)
                    elif credit and not math.isnan(credit):
                        amount = float(credit)
                    else:
                        amount = 0.0

                    description = str(row.get("Description", "")).strip().lower()
                    category = Category.NONE
                    title = description
                # Discover.csv headers:
                #   Trans. Date, Post Date, Description, Amount, Category
                #   Category column present. Examples: Home Improvement, Payments and Credits, Merchandise, Services, Gasoline, Restaurants, Supermarkets.
                case Bank.DISCOVER:
                    try:
                        date = datetime.datetime.strptime(str(row.get("Trans. Date", "")), "%m/%d/%Y").date()
                    except Exception:
                        date = None
                        
                    amount = -1.0 * float(row.get("Amount", 0.0))
                    description = str(row.get("Description", "")).strip().lower()
                    category = Category.NONE
                    title = description
                # CapitalOne.csv headers:
                #   Transaction Date, Posted Date, Card No., Description, Category, Debit, Credit
                #   Category column present. Examples: Gas/Automotive, Merchandise.
                case Bank.CAPITAL_ONE:
                    try:
                        date = datetime.datetime.strptime(str(row.get("Transaction Date", "")), "%Y-%m-%d").date()
                    except Exception:
                        date = None

                    debit = row.get("Debit")
                    credit = row.get("Credit")
                    if debit and not math.isnan(debit):
                        amount = -float(debit)
                    elif credit and not math.isnan(credit):
                        amount = float(credit)
                    else:
                        amount = 0.0

                    description = str(row.get("Description", "")).strip().lower()
                    category = Category.NONE
                    title = description
                # Citi.csv headers:
                #   Status, Date, Description, Debit, Credit
                #   No explicit category column. All transactions default to Category.NONE.
                case Bank.CITI:
                    try:
                        date = datetime.datetime.strptime(str(row.get("Date", "")), "%m/%d/%Y").date()
                    except Exception:
                        date = None

                    debit = row.get("Debit")
                    credit = row.get("Credit")
                    if debit and not math.isnan(debit):
                        amount = -float(debit)
                    elif credit and not math.isnan(credit):
                        amount = float(credit)
                    else:
                        amount = 0.0

                    description = str(row.get("Description", "")).strip().lower()
                    category = Category.NONE
                    title = description
                # Venmo.csv headers:
                #   ID, Datetime, Type, Status, Note, From, To, Amount (total), ...
                #   We want: Datetime, Note, From, To, Amount (total)
                case Bank.VENMO:
                    try:
                        date_str = str(row.get("Datetime", ""))
                        date = datetime.datetime.strptime(date_str[:10], "%Y-%m-%d").date() if date_str else None
                    except Exception:
                        date = None
   
                    amount_raw = str(row.get("Amount (total)", "")).replace("$", "").replace(",", "").strip()
                    # Determine sign from original string
                    if str(row.get("Amount (total)", "")).startswith("-"):
                        amount = -float(amount_raw.replace("-", "")) if amount_raw else 0.0
                    else:
                        amount = float(amount_raw.replace("+", "")) if amount_raw else 0.0

                    note = row.get("Note", "")
                    from_person = str(row.get("From", "")).strip()
                    to_person = str(row.get("To", "")).strip()
                    description = note if note and isinstance(note, str) else f"Venmo: {from_person} -> {to_person}"
                    description = description.strip().lower()

                    category = Category.NONE
                    title = description
                # Fallback: Attempts to use generic columns, including Category and Title if present.
                case _:
                    try:
                        date = datetime.datetime.strptime(str(row.get("Date", "")), "%m/%d/%Y").date()
                    except Exception:
                        date = None
                    amount = float(row.get("Amount", 0.0))
                    description = str(row.get("Description", "")).strip().lower()
                    category = Category.NONE
                    title = str(row.get("Title", "")) if "Title" in row else ""

            # If category is NONE, try to look up by description
            if category == Category.NONE and desc_cat_map:
                # Substring match: find first key in desc_cat_map that is contained in description
                match_found = False
                for dict_key in desc_cat_map:
                    if dict_key in description:
                        title, category = desc_cat_map[dict_key]
                        match_found = True
                        break
                if not match_found:
                    # Use callback to get user input for uncategorized transaction
                    new_description, new_title, new_category = self.get_category_callback(description, amount, str(date))

                    title = new_title.strip()
                    category = next((c for c in Category if c.value == new_category), Category.NONE)

                    # Save new mapping
                    desc_cat_map[new_description.strip().lower()] = (title, category)
                    
                    # Append to CSV file, always ensure a new line
                    with open(DescriptionToCategoryPath, 'a', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile, delimiter=',', lineterminator='\n')
                        writer.writerow([new_description.strip().lower(), title.strip(), category.value])

            transaction = Transaction(
                date=date if date else datetime.date.today(),
                amount=amount,
                description=description,
                bank=bank,
                category=category,
                title=title
            )
            transactions.append(transaction)
        self.logger(f"Parsed {len(transactions)} transactions for {bank.value} from {os.path.basename(file_path)}")
        return Statement(bank=bank, transactions=transactions)



if __name__ == "__main__":
    def logger(msg: str) -> None:
        print(f"[LOG] {msg}")

    def cli_category_callback(description: str, amount: float, date: str) -> tuple[str, str, str]:
        print(f"Uncategorized transaction: '{description}' with amount {amount} on {date}.")
        new_description = input("Description: ").strip().lower()
        new_title = input("Title: ").strip()
        new_category = input("Category: ").strip()
        return new_description, new_title, new_category

    parser: ParseBankStatements = ParseBankStatements(logger, cli_category_callback)
    bank_data: Dict[Bank, str] = parser.find_bank_statements()
    for bank, path in bank_data.items():
        statement: Statement = parser.statement_from_file(bank, path)

