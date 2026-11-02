import datetime
from .DataTypes import Bank, Category, CategoryValue

 # class to represent a single transaction
class Transaction:

	# =============================
	# Constructor
	# =============================
	def __init__(self, date: datetime.date, amount: float, description: str, bank: Bank, category: CategoryValue = Category.NONE, title: str = ""):
		self._date: datetime.date = date
		self._amount: float = amount
		self._description: str = description
		self._title: str = title
		self._category: CategoryValue = category
		self._bank: Bank = bank

	# =============================
	# Transaction Methods
	# =============================
	def is_purchase(self) -> bool:
		"""Return True if the transaction is a purchase."""
		return self._category not in (Category.INCOME, Category.IGNORE, Category.NONE)

	def is_income(self) -> bool:
		"""Return True if the transaction is income."""
		return self._category == Category.INCOME

	# =============================
	# Dunder Methods
	# =============================
	def __repr__(self) -> str:
		return (f"Transaction(date={self._date!r}, amount={self._amount!r}, "
				f"description={self._description!r}, category={self._category!r}, bank={self._bank!r}, title={self._title!r})")

	# =============================
	# Getter/Setter Properties
	# =============================
	@property
	def date(self) -> datetime.date:
		return self._date

	@date.setter
	def date(self, value: datetime.date) -> None:
		self._date = value

	@property
	def title(self) -> str:
		return self._title

	@title.setter
	def title(self, value: str) -> None:
		self._title = value

	@property
	def amount(self) -> float:
		return self._amount

	@amount.setter
	def amount(self, value: float) -> None:
		self._amount = value

	@property
	def description(self) -> str:
		return self._description

	@description.setter
	def description(self, value: str) -> None:
		self._description = value

	@property
	def category(self) -> CategoryValue:
		return self._category

	@category.setter
	def category(self, value: CategoryValue) -> None:
		self._category = value

	@property
	def bank(self) -> Bank:
		return self._bank

	@bank.setter
	def bank(self, value: Bank) -> None:
		self._bank = value
