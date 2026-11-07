"""R27 TEST_006: optional routine confirmations and DB path layout."""
import ast
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)
WINDOW = next(n for n in TREE.body if isinstance(n, ast.ClassDef) and n.name == "MainWindow")


def method(name):
    return next(n for n in WINDOW.body if isinstance(n, ast.FunctionDef) and n.name == name)


class Settings:
    def __init__(self, values=None):
        self.values = dict(values or {})

    def value(self, name, default=None):
        return self.values.get(name, default)


class TestConfirmationsAndPath(unittest.TestCase):
    def test_one_line_buttons_and_breakable_path(self):
        self.assertIn('QPushButton("Create DB")', SOURCE)
        self.assertIn('QPushButton("Open DB")', SOURCE)
        self.assertIn('display_folder.replace', SOURCE)
        self.assertIn('self.mdb_path_label.setText(display_folder + "\\n" + mdb.path.name)', SOURCE)
        self.assertIn('self.mdb_path_label.setToolTip(full_mdb_path)', SOURCE)

    def test_saved_independent_preferences(self):
        opts = ast.unparse(method("_build_menus"))
        for name, key in (
            ("Confirm before Run Scan", "scan/confirm_scan"),
            ("Confirm before Preview Routing", "scan/confirm_preview"),
            ("Confirm before Run Routing", "scan/confirm_move"),
        ):
            self.assertIn(name, opts)
            self.assertIn(key, opts)
        scan = method("_start_scan_operation")
        nested = next(x for x in ast.walk(scan)
                      if isinstance(x, ast.FunctionDef) and x.name == "ask_confirmation")
        ns = {}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[nested], type_ignores=[])),
                     str(ROOT / "app.py"), "exec"), ns)
        fn = ns["ask_confirmation"]
        for raw, expected in ((True, True), (False, False), ("false", False), ("1", True)):
            ns["self"] = SimpleNamespace(settings=Settings({"scan/confirm_preview": raw}))
            self.assertEqual(fn("scan/confirm_preview", True), expected)
        ns["self"] = SimpleNamespace(settings=Settings())
        self.assertFalse(fn("scan/confirm_scan", False))
        self.assertTrue(fn("scan/confirm_preview", True))
        self.assertTrue(fn("scan/confirm_move", True))

    def test_routing_guards_retained(self):
        scan = ast.unparse(method("_start_scan_operation"))
        self.assertIn("_routing_scan_state_decision", scan)
        self.assertIn("routing_preflight_confirmed", scan)
        self.assertIn("if sqlite_mode and real_routing", scan)
        self.assertIn("self.real_routing_check", SOURCE)
        self.assertIn("acknowledged", SOURCE)


if __name__ == "__main__":
    unittest.main()
