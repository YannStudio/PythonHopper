"""Small shared progress event types for long-running Filehopper actions."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class ProgressEvent:
    """Describes a user-visible progress update."""

    phase: str
    message: str
    done: int = 0
    total: int = 0
    percent: int | None = None

    def normalized_percent(self) -> int:
        if self.percent is not None:
            return max(0, min(100, int(self.percent)))
        if self.total > 0:
            return max(0, min(100, int((self.done / self.total) * 100)))
        return 0

    def display_text(self) -> str:
        percent = self.normalized_percent()
        return f"{self.message} ({percent}%)" if percent else self.message


ProgressCallback = Callable[[ProgressEvent], None]
