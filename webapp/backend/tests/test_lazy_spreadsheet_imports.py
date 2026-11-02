import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parents[1]

# Runs in a fresh interpreter because the test suite itself imports pandas.
PROBE = """
import sys
import app.main
loaded = sorted(name for name in ("pandas", "numpy", "openpyxl", "xlrd") if name in sys.modules)
print("AFTER_STARTUP=" + ",".join(loaded))
from app.services import tabular
frame = tabular.read_table("sample.csv", b"Name,Value\\nSynthetic,1\\n")
print("ROWS=" + str(len(frame)))
print("PANDAS_AFTER_READ=" + str("pandas" in sys.modules))
"""


class LazySpreadsheetImportTests(unittest.TestCase):
    def test_spreadsheet_libraries_load_on_first_use_not_at_startup(self):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, PIANIST_SCHEDULING_DB_PATH=str(Path(directory) / "probe.db"))
            env.pop("SCHEDULING_MAESTRO_STARTUP_PROFILE", None)
            env["PYTHONPATH"] = str(BACKEND_DIR)
            result = subprocess.run(
                [sys.executable, "-c", PROBE],
                cwd=directory,
                env=env,
                capture_output=True,
                text=True,
                timeout=120,
            )
        self.assertEqual(result.returncode, 0, result.stderr)
        values = dict(line.split("=", 1) for line in result.stdout.splitlines() if "=" in line)
        self.assertEqual(values["AFTER_STARTUP"], "")
        self.assertEqual(values["ROWS"], "1")
        self.assertEqual(values["PANDAS_AFTER_READ"], "True")


if __name__ == "__main__":
    unittest.main()
