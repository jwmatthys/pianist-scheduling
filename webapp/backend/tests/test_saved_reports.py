import tempfile
import unittest
from pathlib import Path

from fastapi import HTTPException
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models, schemas
from app.database import migrate_database
from app.routers import reports


def payload(markdown="# Report", stamp="3", rows=None):
    return schemas.SavedReportIn(
        markdown=markdown,
        page_break_headings=["Second"],
        table=schemas.SavedReportTable(
            name="Synthetic Report",
            columns=["Student", "Hours", "Needs pianist?"],
            rows=rows if rows is not None else [["Synthetic Student", 1.5, True], ["Other", 0, None]],
        ),
        source_stamp=stamp,
    )


class SavedReportTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(f"sqlite:///{Path(self.temporary_directory.name) / 'session.sqlite3'}")
        migrate_database(self.engine)

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def test_no_report_until_saved(self):
        with Session(self.engine) as db:
            self.assertIsNone(reports.get_saved_report("accompanists", db))

    def test_round_trips_report_and_table_with_types(self):
        with Session(self.engine) as db:
            saved = reports.put_saved_report(payload(), "accompanists", db)
        with Session(self.engine) as db:
            loaded = reports.get_saved_report("accompanists", db)
        self.assertEqual(loaded.markdown, "# Report")
        self.assertEqual(loaded.page_break_headings, ["Second"])
        self.assertEqual(loaded.source_stamp, "3")
        self.assertEqual(loaded.table.rows, [["Synthetic Student", 1.5, True], ["Other", 0, None]])
        self.assertIs(loaded.table.rows[0][2], True)
        self.assertEqual(loaded.generated_at, saved.generated_at)

    def test_saving_again_replaces_only_that_modules_report(self):
        with Session(self.engine) as db:
            reports.put_saved_report(payload("# Old"), "accompanists", db)
            reports.put_saved_report(payload("# Jury", stamp="j"), "jury", db)
            reports.put_saved_report(payload("# New", stamp="9"), "accompanists", db)
            self.assertEqual(db.query(module_models.SavedReport).count(), 2)
            self.assertEqual(reports.get_saved_report("accompanists", db).markdown, "# New")
            self.assertEqual(reports.get_saved_report("accompanists", db).source_stamp, "9")
            self.assertEqual(reports.get_saved_report("jury", db).markdown, "# Jury")

    def test_rejects_oversized_report(self):
        with Session(self.engine) as db:
            with self.assertRaises(HTTPException) as caught:
                reports.put_saved_report(payload("x" * (reports.MAX_SAVED_REPORT_BYTES + 1)), "jury", db)
            self.assertEqual(caught.exception.status_code, 413)

    def test_report_follows_session_uuid_rekeying(self):
        with Session(self.engine) as db:
            reports.put_saved_report(payload(), "clinical", db)
            session = db.query(models.SchedulingSession).one()
            row = db.query(module_models.SavedReport).one()
            self.assertEqual(row.session_uuid, session.session_uuid)


if __name__ == "__main__":
    unittest.main()
