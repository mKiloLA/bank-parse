from enum import Enum
from typing import Union


class Bank(Enum):
    CAPITAL_ONE = "CapitalOne"
    CITI = "Citi"
    COMMERCE = "Commerce"
    DISCOVER = "Discover"
    VENMO = "Venmo"
    NONE = "None"


class Category(Enum):
    MORTGAGE = "Mortgage"
    UTILITIES = "Utilities"
    GIVING = "Giving"
    SUPERMARKET = "Supermarket"
    HOME_IMPROVEMENT = "Home Improvement"
    SERVICES = "Services"
    AUTOMOTIVE = "Automotive"
    ENTERTAINMENT = "Entertainment"
    DINING = "Dining"
    DISCRETIONARY = "Discretionary"
    OTHER = "Other"
    SAVINGS = "Savings"
    INCOME = "Income"
    IGNORE = "Ignore"
    NONE = "None"


CategoryValue = Union[Category, str]


def category_name(category: CategoryValue) -> str:
    return category.value if isinstance(category, Category) else category


def category_from_name(name: str) -> CategoryValue:
    try:
        return Category(name)
    except ValueError:
        return name
