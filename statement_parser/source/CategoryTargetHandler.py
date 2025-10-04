import csv
import os
from typing import Dict, Optional

from .data.DataTypes import Category
from .Utilities import get_appdata_path

class CategoryTargetHandler:
	def __init__(self, csv_path: Optional[str] = None):
		if csv_path is None:
			csv_path = get_appdata_path('categoryTarget.csv')
		self._csv_path = csv_path
		self._targets: Dict[str, float] = {}
		self._read_targets()

	def _read_targets(self) -> None:
		if not os.path.isfile(self._csv_path):
			with open(self._csv_path, 'w', newline='', encoding='utf-8') as f:
				writer = csv.writer(f)
				writer.writerow(['category', 'amount'])
				for cat in Category:
					if cat != Category.NONE:
						writer.writerow([cat.value, 0])
		with open(self._csv_path, newline='', encoding='utf-8') as f:
			reader = csv.DictReader(f)
			for row in reader:
				category = row['category']
				try:
					amount = float(row['amount'])
				except (ValueError, TypeError):
					amount = 0.0
				self._targets[category] = amount

	def get_target(self, category: str) -> float:
		return self._targets.get(category, 0.0)

	def set_target(self, category: str, amount: float) -> None:
		self._targets[category] = amount
		self._write_targets()

	def _write_targets(self) -> None:
		with open(self._csv_path, 'w', newline='', encoding='utf-8') as f:
			writer = csv.writer(f)
			writer.writerow(['category', 'amount'])
			for cat, amt in self._targets.items():
				writer.writerow([cat, amt])

	@property
	def targets(self) -> Dict[str, float]:
		return self._targets.copy()
