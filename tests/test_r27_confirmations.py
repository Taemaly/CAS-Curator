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
    def test_one_line_buttons_and_elided_path(self):
        self.assertIn('QPushButton("Create DB")', SOURCE)
        self.assertIn('QPushButton("Open DB")', SOURCE)
        self.assertIn('class ElidedPathLabel(QLabel):', SOURCE)
        self.assertIn('self.mdb_path_label.setPath(full_mdb_path)', SOURCE)
        self.assertIn('Qt.ElideRight', SOURCE)
        self.assertIn('self.mdb_path_label.setText(display_folder + "\\n" + mdb.path.name)', SOURCE)
        self.assertIn('self.mdb_path_label.setToolTip(full_mdb_path)', SOURCE)

    def test_saved_independent_preferences(self):
        opts = ast.unparse(method("_build_menus"))
        self.assertIn("Confirm before Run Routing", opts)
        self.assertIn("scan/confirm_move", opts)
        self.assertNotIn("Confirm before Run Scan", opts)
        self.assertNotIn("Confirm before Preview Routing", opts)
        self.assertNotIn("scan/confirm_scan", opts)
        self.assertNotIn("scan/confirm_preview", opts)
        scan = method("_start_scan_operation")
        nested = next(x for x in ast.walk(scan)
                      if isinstance(x, ast.FunctionDef) and x.name == "ask_confirmation")
        ns = {}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[nested], type_ignores=[])),
                     str(ROOT / "app.py"), "exec"), ns)
        fn = ns["ask_confirmation"]
        for raw, expected in ((True, True), (False, False), ("false", False), ("1", True)):
            ns["self"] = SimpleNamespace(settings=Settings({"scan/confirm_move": raw}))
            self.assertEqual(fn("scan/confirm_move", True), expected)
        ns["self"] = SimpleNamespace(settings=Settings())
        self.assertTrue(fn("scan/confirm_move", True))
        self.assertIn("confirm_needed = real_routing and", ast.unparse(scan))

    def test_routing_guards_retained(self):
        scan = ast.unparse(method("_start_scan_operation"))
        self.assertIn("_routing_scan_state_decision", scan)
        self.assertIn("routing_preflight_confirmed", scan)
        # Both staged versions fail closed for SQLite moves; the older
        # GitHub GUI also blocks preview until its public integration lands.
        self.assertTrue(
            "if sqlite_mode and real_routing" in scan
            or "if sqlite_mode and (preview_routing or real_routing)" in scan
        )
        self.assertIn("self.real_routing_check", SOURCE)
        self.assertIn("acknowledged", SOURCE)


if __name__ == "__main__":
    unittest.main()
