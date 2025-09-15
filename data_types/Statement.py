from typing import List, Optional
from .DataTypes import Bank
from .Transaction import Transaction

class Statement:
	def __init__(self, bank: Bank, transactions: Optional[List[Transaction]] = None):
		self._bank: Bank = bank
		self._transactions: List[Transaction] = transactions if transactions is not None else []

	@property
	def bank(self) -> Bank:
		return self._bank

	@bank.setter
	def bank(self, value: Bank) -> None:
		self._bank = value

	@property
	def transactions(self) -> List[Transaction]:
		return self._transactions

	@transactions.setter
	def transactions(self, value: List[Transaction]) -> None:
		self._transactions = value

	def add_transaction(self, transaction: Transaction) -> None:
		self._transactions.append(transaction)

	def __repr__(self) -> str:
		return f"Statement(bank={self._bank!r}, transactions={self._transactions!r})"
