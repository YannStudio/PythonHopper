from __future__ import annotations
import sys
from typing import List, Optional

from cli import (
    build_parser,
    cli_suppliers,
    cli_clients,
    cli_bom_check,
    cli_copy,
    cli_copy_per_prod,
    cli_delivery_addresses,
)


def main(argv: Optional[List[str]] = None) -> int:
    """Entry point for the Filehopper application."""
    argv = argv if argv is not None else sys.argv[1:]
    parser = build_parser()
    args = parser.parse_args(argv)

    if getattr(args, "run_tests", False):
        from tests.self_test import run_tests
        return run_tests()

    if not getattr(args, "cmd", None):
        try:
            import tkinter  # noqa: F401
        except Exception:
            parser.print_help()
            return 0

        try:
            from gui import start_gui
        except RuntimeError as exc:
            print(exc, file=sys.stderr)
            return 1

        try:
            start_gui()
        except RuntimeError as exc:
            print(exc, file=sys.stderr)
            return 1
        return 0

    if args.cmd == "suppliers":
        return cli_suppliers(args)
    if args.cmd == "clients":
        return cli_clients(args)
    if args.cmd == "delivery-addresses":
        return cli_delivery_addresses(args)
    if args.cmd == "bom":
        return cli_bom_check(args)
    if args.cmd == "copy":
        return cli_copy(args)
    if args.cmd == "copy-per-prod":
        return cli_copy_per_prod(args)
    parser.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
