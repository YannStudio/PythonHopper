import json
import os
from typing import Any, Dict, Generic, List, TypeVar

T = TypeVar("T")


class BaseDB(Generic[T]):
    """A lightweight base class for simple JSON-backed DBs used in tests.

    Subclasses must implement `schema_version`, `to_dict` and `from_dict`.
    """

    def __init__(self, items: List[T] | None = None):
        self.items: List[T] = list(items) if items else []

    # -- Subclass API --
    def schema_version(self) -> str:
        raise NotImplementedError()

    def to_dict(self, item: T) -> Dict[str, Any]:
        raise NotImplementedError()

    def from_dict(self, d: Dict[str, Any]) -> T:
        raise NotImplementedError()

    # -- Persistence --
    def save(self, path: str) -> None:
        # ensure parent dirs
        parent = os.path.dirname(path)
        if parent and not os.path.exists(parent):
            os.makedirs(parent, exist_ok=True)

        data = {"schema": self.schema_version(), "items": [self.to_dict(i) for i in self.items]}
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)

    def load(self, path: str) -> None:
        if not os.path.exists(path):
            self.items = []
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            # corrupted file or unreadable
            self.items = []
            return

        # Support old format: raw list of items
        if isinstance(data, list):
            raw_items = data
        elif isinstance(data, dict) and "items" in data and isinstance(data["items"], list):
            raw_items = data["items"]
        else:
            # unknown format -> empty
            self.items = []
            return

        loaded: List[T] = []
        for entry in raw_items:
            try:
                loaded.append(self.from_dict(entry))
            except Exception:
                # skip bad entries
                continue
        self.items = loaded
