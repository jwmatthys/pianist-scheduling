import tempfile
import unittest
from datetime import date
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import module_models
from app.database import migrate_database
from app.jury_schemas import JuryPanelFields
from app.services import jury_panel_importer as importer
from app.services.jury import create_panel, list_panels
from app.services.module_lifecycle import active_session_uuid

HEADER = (
    "Schedule Date,Panel Name,Room,Earliest Start,Preferred Start,Jury Length,"
    "Break Needed,Break Every X Juries,Break Length,Meal Break Needed,Meal Start,Meal End\n"
)


class JuryPanelImportTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(f"sqlite:///{Path(self.directory.name) / 's.sqlite3'}")
        migrate_database(self.engine)

    def tearDown(self):
        self.engine.dispose()
        self.directory.cleanup()

    def mapping(self, csv: str) -> importer.JuryPanelImportMapping:
        inspection = importer.inspect_file("panels.csv", csv.encode(), None)
        return importer.JuryPanelImportMapping(**inspection.suggested_mapping)

    def test_import_replaces_panels_and_clears_assignments(self):
        csv = HEADER + (
            "2026-12-01,Voice,Room A,9:00 AM,9:30 AM,20,Yes,4,10,Yes,12:00,13:00\n"
            "12/02/2026,Piano,,09:00,,15,No,,,No,,\n"
        )
        with Session(self.engine) as db:
            old = create_panel(db, JuryPanelFields(
                panel_name="Old", jury_date=date(2026, 11, 1),
                earliest_start_minute=540, jury_length_minutes=20,
            ))
            session_uuid = active_session_uuid(db)
            db.add(module_models.JuryLessonEntry(
                session_uuid=session_uuid,
                source_lesson_uuid="00000000-0000-0000-0000-000000000001",
                student_person_uuid="00000000-0000-0000-0000-000000000002",
                panel_uuid=str(old.panel_uuid),
            ))
            db.flush()

            result = importer.apply_import(db, "panels.csv", csv.encode(), None, self.mapping(csv))
            db.commit()

            self.assertEqual((result.panels_removed, result.assignments_cleared, result.panels_created), (1, 1, 2))
            panels = {panel.panel_name: panel for panel in list_panels(db)}
            self.assertEqual(set(panels), {"Voice", "Piano"})
            self.assertEqual(panels["Voice"].preferred_start_minute, 570)
            self.assertEqual(panels["Voice"].break_every_x_juries, 4)
            self.assertEqual((panels["Voice"].meal_start_minute, panels["Voice"].meal_end_minute), (720, 780))
            self.assertEqual(panels["Piano"].preferred_start_minute, 540)
            self.assertIsNone(panels["Piano"].break_every_x_juries)
            self.assertEqual(panels["Piano"].jury_date, date(2026, 12, 2))
            entry = db.query(module_models.JuryLessonEntry).one()
            self.assertIsNone(entry.panel_uuid)

    def test_invalid_rows_change_nothing(self):
        csv = HEADER + (
            "2026-12-01,Voice,,9:00,,20,No,,,No,,\n"
            "not-a-date,Piano,,9:00,,20,No,,,No,,\n"
            "2026-12-01,voice,,9:00,,20,No,,,No,,\n"
        )
        with Session(self.engine) as db:
            create_panel(db, JuryPanelFields(
                panel_name="Keep", jury_date=date(2026, 11, 1),
                earliest_start_minute=540, jury_length_minutes=20,
            ))
            with self.assertRaises(importer.JuryPanelImportError) as raised:
                importer.apply_import(db, "panels.csv", csv.encode(), None, self.mapping(csv))
            self.assertEqual(len(raised.exception.issues), 2)
            self.assertEqual([panel.panel_name for panel in list_panels(db)], ["Keep"])

    def test_required_mapping_is_enforced(self):
        csv = HEADER + "2026-12-01,Voice,,9:00,,20,No,,,No,,\n"
        mapping = self.mapping(csv)
        mapping.jury_length = None
        with Session(self.engine) as db:
            with self.assertRaises(importer.JuryPanelImportError):
                importer.apply_import(db, "panels.csv", csv.encode(), None, mapping)


if __name__ == "__main__":
    unittest.main()
