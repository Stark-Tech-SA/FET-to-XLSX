import builtins

import pytest

from fet_to_xlsx.ui import app


def test_launcher_falls_back_to_tkinter_when_pyside_is_missing(monkeypatch):
    original_import = builtins.__import__
    called = {}

    def fake_import(name, *args, **kwargs):
        if name == "fet_to_xlsx.ui.main_window":
            raise ModuleNotFoundError("No module named 'PySide6'", name="PySide6")
        return original_import(name, *args, **kwargs)

    def fake_tk_run_app():
        called["tk"] = True
        return 0

    monkeypatch.setattr(builtins, "__import__", fake_import)
    monkeypatch.setattr("fet_to_xlsx.ui.tk_main_window.run_app", fake_tk_run_app)

    assert app.run_app() == 0
    assert called["tk"] is True


def test_launcher_reraises_non_pyside_import_errors(monkeypatch):
    original_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "fet_to_xlsx.ui.main_window":
            raise ModuleNotFoundError("No module named 'other_dependency'", name="other_dependency")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)

    with pytest.raises(ModuleNotFoundError):
        app.run_app()
