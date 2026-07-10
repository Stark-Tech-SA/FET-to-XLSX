"""UI launcher with PySide6 preference and Tkinter fallback."""
from __future__ import annotations

import logging

LOGGER = logging.getLogger(__name__)


def run_app() -> int:
    """Run the best available graphical interface.

    PySide6 is preferred for a modern UI, but the application falls back to
    Tkinter so users can still open it before installing optional UI packages.
    """
    try:
        from fet_to_xlsx.ui.main_window import run_app as run_pyside_app
    except ModuleNotFoundError as exc:
        if exc.name != "PySide6":
            raise
        LOGGER.warning("PySide6 is not installed; using Tkinter fallback UI.")
        from fet_to_xlsx.ui.tk_main_window import run_app as run_tk_app

        return run_tk_app()
    return run_pyside_app()
