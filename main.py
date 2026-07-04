"""Application entry point."""
from __future__ import annotations

import logging

from fet_to_xlsx.ui.main_window import run_app


def main() -> int:
    """Configure logging and launch the desktop application."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
