from typing import Dict, Optional

from .CategoryRuleStore import CategoryRuleStore, CategoryTarget


class CategoryTargetHandler:
    def __init__(self, database_path: Optional[str] = None):
        self._store = CategoryRuleStore(database_path=database_path)

    def get_target(self, category: str) -> float:
        return self._store.get_category_targets().get(category, 0.0)

    def set_target(self, category: str, amount: float) -> None:
        self._store.set_category_target(category, amount)

    def save_category(
        self, name: str, amount: float, original_name: Optional[str] = None
    ) -> None:
        self._store.save_category(name, amount, original_name)

    def delete_category(self, name: str) -> None:
        self._store.delete_category(name)

    @property
    def targets(self) -> Dict[str, float]:
        return self._store.get_category_targets()

    @property
    def category_details(self) -> list[CategoryTarget]:
        return self._store.list_category_targets()
