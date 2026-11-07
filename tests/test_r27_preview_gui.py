"""Regression: GUI must not lose the first headerless SQLite preview route."""
import ast
import tempfile
import unittest
from pathlib import Path
from routing_preview_parser import read_preview_rows


class Item:
    def __init__(self, value): self.value = value
    def setToolTip(self, value): pass
    def setData(self, key, value): pass


class Table:
    def __init__(self): self.rows = []
    def rowCount(self): return len(self.rows)
    def setRowCount(self, count): self.rows = [[None] * 5 for _ in range(count)]
    def insertRow(self, index): self.rows.insert(index, [None] * 5)
    def setItem(self, row, col, value): self.rows[row][col] = value


class Log:
    def __init__(self): self.messages = []
    def appendPlainText(self, value): self.messages.append(value)


def gui_preview_method():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    tree = ast.parse(app_path.read_text(encoding="utf-8"))
    methods = [node for node in ast.walk(tree)
               if isinstance(node, ast.FunctionDef)
               and node.name == "_load_routing_preview_file_if_needed"]
    if len(methods) != 1:
        raise AssertionError("Public GUI preview method missing")
    module = ast.fix_missing_locations(ast.Module(body=[methods[0]], type_ignores=[]))
    namespace = {"Path": Path, "QTableWidgetItem": Item,
                 "Qt": type("Qt", (), {"UserRole": 32})}
    exec(compile(module, str(app_path), "exec"), namespace)
    return namespace["_load_routing_preview_file_if_needed"]


class PreviewGUIRegression(unittest.TestCase):
    def verify(self, report_content, expected_names):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "Routing_Preview.txt"
            path.write_text(report_content, encoding="utf-8")
            instance = type("TestGUI", (), {})()
            instance.last_routing_preview_path = path
            instance.routing_preview_table = Table()
            instance.scan_log = Log()
            instance._routing_preview_display_path = lambda value: value
            gui_preview_method()(instance)
            actual = [row[0].value for row in instance.routing_preview_table.rows]
            self.assertEqual(actual, expected_names)
            self.assertEqual(len(read_preview_rows(path)), len(expected_names))
            self.assertTrue(any(
                "loaded %s routing preview row(s)" % len(expected_names) in message
                for message in instance.scan_log.messages
            ))

    def test_all_four_headerless_sqlite_routes_visible(self):
        lines = "".join(
            "CASP_OBSERVED\\Body\tNo\t/source/p%d.package\t/target/p%d.package\tREADY\n" % (i, i)
            for i in range(4)
        )
        self.verify(lines, ["p0.package", "p1.package", "p2.package", "p3.package"])

    def test_legacy_header_is_skipped(self):
        self.verify(
            "Category\tDuplicate\tSource\tDestination\tNotes\n"
            "EA_OVERRIDE\tNo\t/source/old.package\t/target/old.package\tREADY\n",
            ["old.package"],
        )


if __name__ == "__main__":
    unittest.main()
