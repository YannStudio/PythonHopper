"""Backward-compatible public API for order utilities.

The implementation lives in :mod:`orders.file_operations`.  This package keeps
``import orders`` working for the GUI, CLI, and tests while allowing the module
to be split into smaller files over time.
"""

from __future__ import annotations

import sys
from types import ModuleType

from . import file_operations as _file_operations


def _export_from_file_operations() -> None:
    for name in dir(_file_operations):
        if name.startswith("__") and name.endswith("__"):
            continue
        globals()[name] = getattr(_file_operations, name)


_export_from_file_operations()


def __getattr__(name: str) -> object:
    return getattr(_file_operations, name)


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(dir(_file_operations)))


class _OrdersModule(ModuleType):
    def __setattr__(self, name: str, value: object) -> None:
        super().__setattr__(name, value)
        if hasattr(_file_operations, name):
            setattr(_file_operations, name, value)


sys.modules[__name__].__class__ = _OrdersModule
