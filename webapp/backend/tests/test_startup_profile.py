import importlib
import io
import os
import sys
import unittest
from contextlib import redirect_stderr
from unittest import mock

from app import startup_profile


def load_profile(flag):
    startup_profile._uninstall_import_timing()
    env = {startup_profile.ENV_FLAG: flag} if flag is not None else {}
    with mock.patch.dict(os.environ, env, clear=False):
        if flag is None:
            os.environ.pop(startup_profile.ENV_FLAG, None)
        return importlib.reload(startup_profile)


class StartupProfileTests(unittest.TestCase):
    def tearDown(self):
        startup_profile._uninstall_import_timing()
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop(startup_profile.ENV_FLAG, None)
            importlib.reload(startup_profile)

    def test_emit_event_prints_a_fixed_token_only_when_the_shell_requests_it(self):
        buffer = io.StringIO()
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop(startup_profile.EVENTS_ENV_FLAG, None)
            with mock.patch("sys.stdout", buffer):
                startup_profile.emit_event("database-initialization-begin")
        self.assertEqual(buffer.getvalue(), "")
        with mock.patch.dict(os.environ, {startup_profile.EVENTS_ENV_FLAG: "1"}):
            with mock.patch("sys.stdout", buffer):
                startup_profile.emit_event("database-initialization-begin")
        self.assertEqual(buffer.getvalue(), "STARTUP-EVENT database-initialization-begin\n")

    def test_disabled_does_not_patch_import_machinery_or_record_blocks(self):
        import importlib._bootstrap as bootstrap

        original = bootstrap._find_and_load
        profile = load_profile(None)
        self.assertIs(bootstrap._find_and_load, original)
        with profile.measure("Block"):
            pass
        self.assertEqual(profile._blocks, [])
        self.assertEqual(profile._imports, [])

    def test_enabled_records_import_times_and_ranked_report(self):
        import importlib._bootstrap as bootstrap

        profile = load_profile("1")
        sys.modules.pop("colorsys", None)
        importlib.import_module("colorsys")
        with profile.measure("Synthetic block"):
            pass
        self.assertIn("colorsys", [name for name, _cumulative, _self in profile._imports])
        buffer = io.StringIO()
        with redirect_stderr(buffer):
            profile.log_summary()
        text = buffer.getvalue()
        self.assertIn("Top startup contributors", text)
        self.assertIn("Synthetic block", text)
        self.assertIsNone(profile._original_find_and_load)
        self.assertNotEqual(bootstrap._find_and_load.__name__, "timed_find_and_load")

    def test_disabled_by_default_emits_nothing(self):
        profile = load_profile(None)
        buffer = io.StringIO()
        with redirect_stderr(buffer):
            profile.mark("Backend process entry")
            profile.log_summary()
        self.assertFalse(profile.enabled())
        self.assertEqual(buffer.getvalue(), "")

    def test_enabled_logs_each_label_once_and_summarizes_measured_spans(self):
        profile = load_profile("1")
        buffer = io.StringIO()
        with redirect_stderr(buffer):
            profile.mark("Migration check begin")
            profile.mark("Migration check begin")
            profile.mark("Migration check complete")
            profile.log_summary()
        lines = buffer.getvalue().splitlines()
        self.assertEqual(sum("Migration check begin" in line for line in lines), 1)
        self.assertTrue(all(line.startswith("[STARTUP backend") for line in lines))
        self.assertTrue(any("Migration check:" in line for line in lines))
        self.assertFalse(any("Database initialization:" in line for line in lines))


if __name__ == "__main__":
    unittest.main()
