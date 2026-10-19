from typing import TypeVar

T = TypeVar("T")


class IndexedList(list[T]):
    def __init__(self):
        self._index: dict[T, int] = {}

    def get_index(self, item: T) -> int:
        if item not in self._index:
            self._index[item] = len(self)
            self.append(item)
        return self._index[item]

    def get_item(self, index: int) -> T | None:
        return self[index] if 0 <= index < len(self) else None

    def get_items(self) -> list[T]:
        return list(self._index.keys())
