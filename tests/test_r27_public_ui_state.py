"""UI regression checks without requiring PySide6 in the CI environment.

The actual Qt window/Windows scaling still requires a user acceptance test.
"""
import ast
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
TREE = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
WINDOW = next(n for n in TREE.body if isinstance(n, ast.ClassDef) and n.name == "MainWindow")


def method(name, **namespace):
    item = next(n for n in WINDOW.body if isinstance(n, ast.FunctionDef) and n.name == name)
    scope = dict(namespace)
    module = ast.fix_missing_locations(ast.Module(body=[item], type_ignores=[]))
    exec(compile(module, str(ROOT / "app.py"), "exec"), scope)
    return scope[name]


class Settings:
    def __init__(self, values=None): self.data = dict(values or {})
    def value(self, name, fallback=None): return self.data.get(name, fallback)
    def setValue(self, name, value): self.data[name] = value
    def sync(self): pass


class Label:
    def __init__(self): self.messages = []
    def setText(self, value): self.messages.append(value)
    def showMessage(self, value, *args): self.messages.append(value)
    def appendPlainText(self, value): self.messages.append(value)


class Dialogs:
    save_path = None
    notices = []

    @classmethod
    def getSaveFileName(cls, *args):
        return str(cls.save_path), ""

    @classmethod
    def warning(cls, *args): cls.notices.append(("warning", args))

    @classmethod
    def critical(cls, *args): cls.notices.append(("critical", args))


class PublicLibraryUIRegression(unittest.TestCase):
    def test_first_upgrade_prefers_last_sqlite_over_old_mdb(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sqlite = root / "recent.db"
            xlsx = root / "old.xlsx"
            sqlite.touch()
            xlsx.touch()
            opened = []
            gui = SimpleNamespace(
                settings=Settings({"last_library_db": str(sqlite), "last_mdb": str(xlsx)}),
                load_file=opened.append, statusBar=Label
            )
            method("_restore_last_mdb", Path=Path)(gui)
            self.assertEqual(opened, [sqlite])

    def test_unified_latest_library_preference_overrides_legacy_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            chosen = root / "last_opened.xlsx"
            sqlite = root / "older.db"
            chosen.touch()
            sqlite.touch()
            opened = []
            gui = SimpleNamespace(
                settings=Settings({"last_library_path": str(chosen), "last_library_db": str(sqlite)}),
                load_file=opened.append, statusBar=Label
            )
            method("_restore_last_mdb", Path=Path)(gui)
            self.assertEqual(opened, [chosen])

    def test_create_database_is_real_sqlite_and_does_not_overwrite_existing(self):
        from sqlite_store import create_fresh_sqlite_library, validate_sqlite_library

        creator = method("create_new_database", Path=Path, QFileDialog=Dialogs,
                         QMessageBox=Dialogs,
                         create_fresh_sqlite_library=create_fresh_sqlite_library)
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "new.db"
            Dialogs.save_path = target
            Dialogs.notices.clear()
            opened = []
            gui = SimpleNamespace(
                settings=Settings({"last_dir": tmp}), _confirm_pending_changes=lambda: True,
                load_file=opened.append, statusBar=Label, scan_mdb_override=True,
            )
            creator(gui)
            self.assertEqual(opened, [target.resolve()])
            self.assertEqual(validate_sqlite_library(target)["integrity"], "ok")
            self.assertFalse(Dialogs.notices)
            self.assertFalse(gui.scan_mdb_override)

            original = target.read_bytes()
            creator(gui)
            self.assertEqual(target.read_bytes(), original)
            self.assertEqual(opened, [target.resolve()])
            self.assertEqual(Dialogs.notices[-1][0], "warning")

    def test_refresh_completion_log_is_emitted_only_after_library_loaded(self):
        source = method("_on_loaded", MDBData=object)
        # Extract just the end-of-load success branch so Qt can be mocked.
        func = next(n for n in WINDOW.body if isinstance(n, ast.FunctionDef) and n.name == "_on_loaded")
        branch = next(n for n in ast.walk(func)
                      if isinstance(n, ast.If) and "scan_completion_waiting_for_library" in ast.unparse(n.test))
        logs = Label()
        gui = SimpleNamespace(
            _scan_completion_waiting_for_library=True,
            scan_log=logs, scan_status_label=Label(), statusBar=Label,
            _set_scan_progress_complete=lambda: None,
        )
        module = ast.fix_missing_locations(ast.Module(body=[branch], type_ignores=[]))
        exec(compile(module, "refresh_finish", "exec"),
             {"self": gui, "mdb": SimpleNamespace(records=[{}, {}, {}, {}])})
        self.assertFalse(gui._scan_completion_waiting_for_library)
        self.assertTrue(any("refresh COMPLETE" in line and "4 packages" in line
                            for line in logs.messages))

    def test_compact_buttons_and_legacy_excel_workflow_retained(self):
        text = (ROOT / "app.py").read_text(encoding="utf-8")
        for snippet in (
            'QPushButton("Create DB")',
            'QPushButton("Open DB")',
            'self.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)',
            'self.mdb_path_label.setPath(full_mdb_path)',
            'class ElidedPathLabel(QLabel)',
            'self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)',
            'Create legacy Excel MDB…',
            'def create_new_mdb(self):',
            'self.settings.setValue("last_library_path"',
        ):
            self.assertIn(snippet, text, snippet)


if __name__ == "__main__":
    unittest.main()
