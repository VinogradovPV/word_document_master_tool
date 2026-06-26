from __future__ import annotations

import argparse


def build_parser() -> argparse.ArgumentParser:
    """
    Создает парсер аргументов командной строки.
    """
    parser = argparse.ArgumentParser(
        prog="wdmt",
        description="Word Document Master Tool command line interface",
    )
    parser.add_argument(
        "--version",
        action="store_true",
        help="Показать версию пакета и выйти.",
    )
    return parser


def main() -> int:
    """
    Точка входа CLI.
    """
    parser = build_parser()
    args = parser.parse_args()

    if args.version:
        print("word-document-master-tool 0.1.0")
        return 0

    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
