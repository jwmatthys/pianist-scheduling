"""Synthetic-data tests for the Clinical Placement persistence foundation."""

import tempfile
import unittest
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app import clinical_models, models
from app.clinical_models import (
    ClinicalModuleState,
    ClinicalPlacement,
    PlacementType,
    TagState,
    normalize_tag_name,
    tag_identity_key,
)
from app.database import LATEST_SCHEMA_VERSION, migrate_database
from app.services.availability import AvailabilityWindow
from app.services.clinical_repository import (
    ClinicalRepository,
    ClinicalValidationError,
    PlacementOpportunityInput,
)

CLINICAL_TABLES = {
    "clinical_tags",
    "clinical_placement_opportunities",
    "clinical_opportunity_tags",
    "clinical_students",
    "clinical_student_availability_windows",
    "clinical_student_tag_states",
    "clinical_placements",
    "clinical_module_states",
}


def opportunity_input(**overrides) -> PlacementOpportunityInput:
    values = dict(
        site_name="Synthetic Clinic",
        site_address="1 Example Way",
        contact_name="Casey Contact",
        contact_email="contact@example.test",
        contact_phone="555-0100",
        day_of_week="Tue",
        start_minute=8 * 60,
        end_minute=11 * 60,
        session_length_minutes=60,
        travel_minutes=15,
    )
    values.update(overrides)
    return PlacementOpportunityInput(**values)


class ClinicalTestCase(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(
            f"sqlite:///{Path(self.temporary_directory.name) / 'session.sqlite3'}",
            connect_args={"check_same_thread": False},
        )
        migrate_database(self.engine)
        self.db = Session(self.engine)
        self.repo = ClinicalRepository(self.db)

    def tearDown(self):
        self.db.close()
        self.engine.dispose()
        self.temporary_directory.cleanup()


class TagTests(ClinicalTestCase):
    def test_normalization_trims_and_uppercases(self):
        self.assertEqual(normalize_tag_name("  mental health "), "MENTAL HEALTH")
        self.assertEqual(normalize_tag_name(None), "")
        self.assertEqual(tag_identity_key("Virtual"), tag_identity_key("VIRTUAL"))

    def test_created_tag_is_uppercase_with_normalized_name(self):
        tag = self.repo.create_tag("  advanced ")
        self.db.commit()
        self.assertEqual(tag.name, "ADVANCED")
        self.assertEqual(tag.normalized_name, "advanced")

    def test_duplicates_are_rejected_case_insensitively(self):
        self.repo.create_tag("Virtual")
        with self.assertRaises(ClinicalValidationError) as raised:
            self.repo.create_tag("vIRTUAL")
        self.assertEqual(raised.exception.code, "DUPLICATE_TAG")

    def test_get_or_create_reuses_existing_tag(self):
        first = self.repo.get_or_create_tag("virtual")
        second = self.repo.get_or_create_tag("VIRTUAL ")
        self.assertEqual(first.id, second.id)
        self.assertEqual(len(self.repo.list_tags()), 1)

    def test_database_enforces_uniqueness_independently(self):
        self.repo.create_tag("Virtual")
        self.db.commit()
        self.db.add(clinical_models.ClinicalTag(name="VIRTUAL", normalized_name="virtual"))
        with self.assertRaises(IntegrityError):
            self.db.commit()

    def test_empty_tag_is_rejected(self):
        with self.assertRaises(ClinicalValidationError):
            self.repo.create_tag("   ")

    def test_rename_preserves_associations_and_rejects_collisions(self):
        student = self.repo.create_student("S1", "Student One")
        opportunity = self.repo.create_opportunity(opportunity_input(), ["adult"])
        tag = self.repo.find_tag("adult")
        self.repo.set_tag_state(student.id, tag.id, TagState.REQUIRED)
        other = self.repo.create_tag("child")

        renamed = self.repo.rename_tag(tag.id, "adults")
        self.db.commit()

        self.assertEqual(renamed.name, "ADULTS")
        self.assertEqual([t.name for t in opportunity.tags], ["ADULTS"])
        self.assertEqual(self.repo.get_tag_state(student.id, tag.id), TagState.REQUIRED)
        with self.assertRaises(ClinicalValidationError):
            self.repo.rename_tag(other.id, "Adults")
        self.assertEqual(self.repo.rename_tag(tag.id, "ADULTS").id, tag.id)

    def test_unused_tags_remain_registered(self):
        opportunity = self.repo.create_opportunity(opportunity_input(), ["adult"])
        tag = self.repo.find_tag("adult")
        self.assertTrue(self.repo.is_tag_used_by_opportunity(tag.id))
        self.repo.update_opportunity(opportunity.id, opportunity_input(), [])
        self.assertFalse(self.repo.is_tag_used_by_opportunity(tag.id))
        self.assertEqual([t.name for t in self.repo.list_tags()], ["ADULT"])

    def test_delete_removes_clinical_associations_and_reports_affected_students(self):
        tag = self.repo.create_tag("adult")
        first = self.repo.create_student("S1")
        second = self.repo.create_student("S2")
        self.repo.create_student("S3")
        self.repo.set_tag_state(first.id, tag.id, TagState.REQUIRED)
        self.repo.set_tag_state(second.id, tag.id, TagState.RESTRICTED)
        opportunity = self.repo.create_opportunity(opportunity_input(), ["adult"])
        self.assertEqual(self.repo.count_students_with_tag_selection(tag.id), 2)

        affected = self.repo.delete_tag(tag.id)
        self.db.commit()

        self.assertEqual(affected, 2)
        self.assertEqual(self.repo.list_tags(), [])
        self.assertEqual(self.db.query(clinical_models.ClinicalStudentTagState).count(), 0)
        self.db.refresh(opportunity)
        self.assertEqual(opportunity.tags, [])

    def test_neutral_state_is_stored_as_absence(self):
        tag = self.repo.create_tag("adult")
        student = self.repo.create_student("S1")
        self.repo.set_tag_state(student.id, tag.id, TagState.PREFERRED)
        self.repo.set_tag_state(student.id, tag.id, TagState.NEUTRAL)
        self.assertEqual(self.repo.get_tag_state(student.id, tag.id), TagState.NEUTRAL)
        self.assertEqual(self.db.query(clinical_models.ClinicalStudentTagState).count(), 0)


class OpportunityTests(ClinicalTestCase):
    def test_creation_defaults_and_tags(self):
        opportunity = self.repo.create_opportunity(opportunity_input(), ["virtual", "VIRTUAL", "advanced", " "])
        self.db.commit()
        self.db.refresh(opportunity)

        self.assertEqual(opportunity.day_of_week, "Tuesday")
        self.assertEqual(opportunity.capacity, 1)
        self.assertEqual(opportunity.travel_minutes, 15)
        self.assertEqual([t.name for t in opportunity.tags], ["ADVANCED", "VIRTUAL"])
        self.assertTrue(opportunity.id)

    def test_same_site_rows_remain_distinct_opportunities(self):
        first = self.repo.create_opportunity(opportunity_input())
        second = self.repo.create_opportunity(opportunity_input(day_of_week="Thursday"))
        self.assertNotEqual(first.id, second.id)
        self.assertEqual(self.db.query(clinical_models.PlacementOpportunity).count(), 2)

    def test_invalid_values_are_rejected(self):
        invalid = [
            opportunity_input(day_of_week="Funday"),
            opportunity_input(start_minute=600, end_minute=600),
            opportunity_input(session_length_minutes=0),
            opportunity_input(travel_minutes=-1),
            opportunity_input(capacity=0),
        ]
        for values in invalid:
            with self.assertRaises(ClinicalValidationError):
                self.repo.create_opportunity(values)

    def test_database_constraints_reject_invalid_rows(self):
        self.db.add(clinical_models.PlacementOpportunity(
            day_of_week="Monday", start_minute=600, end_minute=540, session_length_minutes=30,
        ))
        with self.assertRaises(IntegrityError):
            self.db.commit()

    def test_deleting_in_use_opportunity_removes_only_its_placements(self):
        doomed = self.repo.create_opportunity(opportunity_input(site_name="Doomed"))
        kept = self.repo.create_opportunity(opportunity_input(site_name="Kept"))
        affected = self.repo.create_student("S1")
        elsewhere = self.repo.create_student("S2")
        unplaced = self.repo.create_student("S3")
        self.repo.set_placement(affected.id, PlacementType.AUTOMATIC, doomed.id, 480)
        self.repo.set_placement(elsewhere.id, PlacementType.MANUAL, kept.id, 490)
        self.repo.set_placement(unplaced.id, PlacementType.MANUALLY_UNPLACED)
        self.assertEqual(self.repo.count_placements_for_opportunity(doomed.id), 1)

        removed = self.repo.delete_opportunity(doomed.id)
        self.db.commit()

        self.assertEqual(removed, 1)
        self.assertIsNone(self.repo.get_placement(affected.id))
        self.assertIsNotNone(self.repo.get_placement(elsewhere.id))
        self.assertEqual(self.repo.get_placement(unplaced.id).placement_type, "manuallyUnplaced")
        self.assertTrue(self.repo.get_state().is_schedule_stale)


class StudentTests(ClinicalTestCase):
    def test_creation_with_recurring_availability(self):
        windows = [
            AvailabilityWindow("Monday", 540, 720, "Available"),
            AvailabilityWindow("Monday", 720, 780, "Tentative", source="import", source_row=3),
            AvailabilityWindow("Friday", 600, 660, "Unavailable"),
        ]
        student = self.repo.create_student("  s-100 ", " Pat Student ", "pat@example.test", windows)
        self.db.commit()
        self.db.refresh(student)

        self.assertEqual((student.student_id, student.name, student.email), ("s-100", "Pat Student", "pat@example.test"))
        self.assertEqual(self.repo.student_availability(student.id), sorted(
            windows, key=lambda w: (w.day, w.start_minute)
        ))
        self.assertEqual(len(student.availability), 3)

    def test_availability_replacement_and_missing_availability(self):
        student = self.repo.create_student("S1")
        self.assertEqual(self.repo.student_availability(student.id), [])
        self.repo.replace_student_availability(student.id, [AvailabilityWindow("Sunday", 0, 60, "Available")])
        self.repo.replace_student_availability(student.id, [AvailabilityWindow("Monday", 60, 120, "Tentative")])
        self.assertEqual(
            [(w.day, w.status) for w in self.repo.student_availability(student.id)], [("Monday", "Tentative")]
        )

    def test_invalid_window_is_rejected_and_state_preserved(self):
        student = self.repo.create_student("S1", availability=[AvailabilityWindow("Monday", 60, 120, "Available")])
        with self.assertRaises(ClinicalValidationError):
            self.repo.replace_student_availability(student.id, [AvailabilityWindow("Monday", 120, 60, "Available")])
        self.assertEqual(len(self.repo.student_availability(student.id)), 1)

    def test_student_id_is_required_and_unique_case_insensitively(self):
        with self.assertRaises(ClinicalValidationError):
            self.repo.create_student("  ")
        self.repo.create_student("abc")
        with self.assertRaises(ClinicalValidationError) as raised:
            self.repo.create_student("ABC")
        self.assertEqual(raised.exception.code, "DUPLICATE_STUDENT_ID")

    def test_delete_removes_dependents(self):
        tag = self.repo.create_tag("adult")
        opportunity = self.repo.create_opportunity(opportunity_input())
        student = self.repo.create_student("S1", availability=[AvailabilityWindow("Monday", 60, 120, "Available")])
        self.repo.set_tag_state(student.id, tag.id, TagState.PREFERRED)
        self.repo.set_placement(student.id, PlacementType.MANUAL, opportunity.id, 490)

        self.repo.delete_student(student.id)
        self.db.commit()

        for model in (
            clinical_models.ClinicalStudent,
            clinical_models.ClinicalStudentAvailabilityWindow,
            clinical_models.ClinicalStudentTagState,
            ClinicalPlacement,
        ):
            self.assertEqual(self.db.query(model).count(), 0)
        self.assertEqual(self.db.query(clinical_models.PlacementOpportunity).count(), 1)

    def test_tag_states_cycle_values_persist(self):
        tag = self.repo.create_tag("adult")
        student = self.repo.create_student("S1")
        for state in (TagState.REQUIRED, TagState.PREFERRED, TagState.RESTRICTED):
            self.repo.set_tag_state(student.id, tag.id, state)
            self.db.commit()
            self.assertEqual(self.repo.get_tag_state(student.id, tag.id), state)


class PlacementTests(ClinicalTestCase):
    def setUp(self):
        super().setUp()
        self.opportunity = self.repo.create_opportunity(opportunity_input())
        self.student = self.repo.create_student("S1")

    def test_each_placement_type_persists(self):
        for placement_type, opportunity_id, start in (
            (PlacementType.AUTOMATIC, self.opportunity.id, 495),
            (PlacementType.MANUAL, self.opportunity.id, 490),
            (PlacementType.MANUALLY_UNPLACED, None, None),
        ):
            self.repo.set_placement(self.student.id, placement_type, opportunity_id, start)
            self.db.commit()
            self.db.expire_all()
            stored = self.repo.get_placement(self.student.id)
            self.assertEqual(
                (stored.placement_type, stored.placement_opportunity_id, stored.start_minute),
                (placement_type.value, opportunity_id, start),
            )
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 1)

    def test_invalid_placement_shapes_are_rejected(self):
        with self.assertRaises(ClinicalValidationError):
            self.repo.set_placement(self.student.id, PlacementType.MANUAL)
        with self.assertRaises(ClinicalValidationError):
            self.repo.set_placement(self.student.id, PlacementType.MANUALLY_UNPLACED, self.opportunity.id, 480)
        with self.assertRaises(ClinicalValidationError):
            self.repo.set_placement(self.student.id, PlacementType.MANUAL, "missing", 480)

    def test_database_rejects_malformed_placement_and_second_placement(self):
        self.db.add(ClinicalPlacement(clinical_student_id=self.student.id, placement_type="manual"))
        with self.assertRaises(IntegrityError):
            self.db.commit()
        self.db.rollback()
        self.db.add(ClinicalPlacement(
            clinical_student_id=self.student.id, placement_type="automatic",
            placement_opportunity_id=self.opportunity.id, start_minute=480,
        ))
        self.db.commit()
        self.db.add(ClinicalPlacement(
            clinical_student_id=self.student.id, placement_type="manuallyUnplaced",
        ))
        with self.assertRaises(IntegrityError):
            self.db.commit()


class ModuleStateTests(ClinicalTestCase):
    def test_state_row_exists_with_defaults_and_persists(self):
        state = self.repo.get_state()
        self.assertEqual(
            (state.schedule_revision, state.source_data_revision, state.is_schedule_stale), (0, 0, False)
        )
        state.schedule_revision = 4
        state.source_data_revision = 7
        state.is_schedule_stale = True
        self.db.commit()
        self.db.expire_all()
        stored = self.db.query(ClinicalModuleState).one()
        self.assertEqual((stored.schedule_revision, stored.source_data_revision, stored.is_schedule_stale), (4, 7, True))

    def test_source_change_marks_stale_only_when_placements_exist(self):
        opportunity = self.repo.create_opportunity(opportunity_input())
        student = self.repo.create_student("S1")
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        revision = self.repo.get_state().source_data_revision

        self.repo.set_placement(student.id, PlacementType.AUTOMATIC, opportunity.id, 480)
        self.repo.update_opportunity(opportunity.id, opportunity_input(capacity=3))

        state = self.repo.get_state()
        self.assertTrue(state.is_schedule_stale)
        self.assertEqual(state.source_data_revision, revision + 1)
        self.assertEqual(self.repo.get_placement(student.id).placement_type, "automatic")

    def test_clear_all_placements_resets_staleness_but_keeps_source_data(self):
        opportunity = self.repo.create_opportunity(opportunity_input())
        student = self.repo.create_student("S1")
        self.repo.set_placement(student.id, PlacementType.MANUAL, opportunity.id, 490)
        self.repo.update_opportunity(opportunity.id, opportunity_input(capacity=2))
        self.assertTrue(self.repo.get_state().is_schedule_stale)

        self.assertEqual(self.repo.clear_all_placements(), 1)

        state = self.repo.get_state()
        self.assertFalse(state.is_schedule_stale)
        self.assertEqual(state.schedule_revision, 0)
        self.assertEqual(self.db.query(clinical_models.PlacementOpportunity).count(), 1)
        self.assertEqual(self.db.query(clinical_models.ClinicalStudent).count(), 1)


class ClinicalMigrationTests(ClinicalTestCase):
    def test_fresh_database_has_clinical_tables_and_state_for_the_session(self):
        tables = {row[0] for row in self.engine.connect().exec_driver_sql(
            "SELECT name FROM sqlite_master WHERE type = 'table'"
        )}
        self.assertTrue(CLINICAL_TABLES <= tables)
        session_uuid = self.db.query(models.SchedulingSession.session_uuid).scalar()
        self.assertEqual(self.db.query(ClinicalModuleState).one().session_uuid, session_uuid)

    def test_v9_database_upgrades_preserving_existing_data(self):
        with Session(self.engine) as db:
            db.add(models.Pianist(id=7, name="Synthetic Pianist"))
            db.commit()
        with self.engine.begin() as connection:
            for table in sorted(CLINICAL_TABLES - {"clinical_module_states"}, reverse=True):
                connection.exec_driver_sql(f"DROP TABLE {table}")
            connection.exec_driver_sql("DROP TABLE clinical_module_states")
            connection.exec_driver_sql("DROP TABLE saved_reports")
            connection.exec_driver_sql("PRAGMA user_version = 9")

        report = migrate_database(self.engine)

        self.assertEqual(report.initial_version, 9)
        self.assertEqual(report.current_version, LATEST_SCHEMA_VERSION)
        self.assertEqual(report.applied_migrations, ("add_clinical_placement_module_tables", "add_saved_reports"))
        with Session(self.engine) as db:
            self.assertEqual(db.get(models.Pianist, 7).name, "Synthetic Pianist")
            state = db.query(ClinicalModuleState).one()
            self.assertEqual(
                state.session_uuid, db.query(models.SchedulingSession.session_uuid).scalar()
            )
            self.assertEqual(state.source_data_revision, 0)

    def test_migration_is_idempotent_at_current_version(self):
        report = migrate_database(self.engine)
        self.assertEqual(report.applied_migrations, ())
        self.assertEqual(len(self.db.query(ClinicalModuleState).all()), 1)

    def test_clinical_schema_has_no_cross_module_foreign_keys(self):
        for table in clinical_models.ModuleBase.metadata.sorted_tables:
            if not table.name.startswith("clinical_"):
                continue
            for foreign_key in table.foreign_keys:
                self.assertTrue(foreign_key.column.table.name.startswith("clinical_"), table.name)


if __name__ == "__main__":
    unittest.main()
