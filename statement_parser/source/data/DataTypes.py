from enum import Enum

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
