"""Application entry point."""
from __future__ import annotations

import argparse
import logging

from fet_to_xlsx.ui.app import run_app


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line parser used by the desktop launcher."""
    parser = argparse.ArgumentParser(description="Conversor FET → Excel")
    parser.add_argument("--version", action="version", version="FET-to-XLSX 0.1.0")
    return parser


def main() -> int:
    """Configure logging and launch the desktop application."""
    build_parser().parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
