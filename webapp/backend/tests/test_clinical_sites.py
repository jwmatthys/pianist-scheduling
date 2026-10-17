"""Synthetic-data tests for Clinical Sites and Tag Manager repository and API behavior."""

import unittest

from fastapi import HTTPException

from app import clinical_schemas as schemas
from app.clinical_models import PlacementType, TagState
from app.routers import clinical
from app.services.clinical_repository import ClinicalValidationError
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input


class TagManagerRepositoryTests(ClinicalTestCase):
    def revision(self):
        return self.repo.get_state().source_data_revision

    def test_create_normalizes_and_rejects_case_insensitive_duplicates(self):
        tag = self.repo.create_tag("  mental health ")
        self.assertEqual(tag.name, "MENTAL HEALTH")
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.create_tag("Mental Health")
        self.assertEqual(caught.exception.code, "DUPLICATE_TAG")

    def test_rename_and_duplicate_rename(self):
        a = self.repo.create_tag("a")
        self.repo.create_tag("b")
        self.assertEqual(self.repo.rename_tag(a.id, " c ").name, "C")
        with self.assertRaises(ClinicalValidationError):
            self.repo.rename_tag(a.id, "B")

    def test_unused_detection_ignores_student_use(self):
        used, unused = self.repo.create_tag("used"), self.repo.create_tag("unused")
        self.repo.create_opportunity(opportunity_input(), tag_ids=[used.id])
        summaries = {s.tag.name: s for s in self.repo.tag_summaries()}
        self.assertFalse(summaries["USED"].is_unused)
        self.assertTrue(summaries["UNUSED"].is_unused)
        self.assertEqual(summaries["USED"].opportunity_count, 1)

    def test_delete_tag_cleans_associations_and_counts_students(self):
        tag = self.repo.create_tag("x")
        opp = self.repo.create_opportunity(opportunity_input(), tag_ids=[tag.id])
        student = self.repo.create_student(student_id="S1", name="Pat Test", email="")
        self.repo.set_tag_state(student.id, tag.id, TagState.REQUIRED)
        self.assertEqual(self.repo.delete_tag(tag.id), 1)
        self.db.refresh(opp)
        self.assertEqual(list(opp.tags), [])
        self.assertEqual(self.repo.list_tags(), [])

    def test_delete_unused_tags_keeps_tags_used_by_sites_or_students(self):
        site_tag, student_tag, unused = (self.repo.create_tag(n) for n in ("site", "student", "unused"))
        self.repo.create_opportunity(opportunity_input(), tag_ids=[site_tag.id])
        student = self.repo.create_student(student_id="S1", name="Pat Test", email="")
        self.repo.set_tag_state(student.id, student_tag.id, TagState.PREFERRED)
        before = self.revision()
        self.assertEqual(self.repo.delete_unused_tags(), 1)
        self.assertEqual(sorted(tag.name for tag in self.repo.list_tags()), ["SITE", "STUDENT"])
        self.assertEqual(self.revision(), before + 1)
        self.assertEqual(self.repo.delete_unused_tags(), 0)
        self.assertEqual(self.revision(), before + 1)

    def test_every_mutation_bumps_revision(self):
        seen = [self.revision()]
        tag = self.repo.create_tag("t"); seen.append(self.revision())
        self.repo.rename_tag(tag.id, "t2"); seen.append(self.revision())
        opp = self.repo.create_opportunity(opportunity_input()); seen.append(self.revision())
        self.repo.update_opportunity(opp.id, opportunity_input(site_name="Other")); seen.append(self.revision())
        self.repo.set_opportunity_tags(opp.id, [tag.id]); seen.append(self.revision())
        self.repo.delete_tag(tag.id); seen.append(self.revision())
        self.repo.delete_opportunity(opp.id); seen.append(self.revision())
        self.assertEqual(seen, list(range(seen[0], seen[0] + len(seen))))

    def test_stale_only_when_placements_exist(self):
        opp = self.repo.create_opportunity(opportunity_input())
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        student = self.repo.create_student(student_id="S1", name="Pat Test", email="")
        self.repo.set_placement(student.id, PlacementType.MANUAL, opp.id, 8 * 60)
        self.repo.create_tag("late")
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_site_name_required_and_basic_validation(self):
        for overrides in (
            {"site_name": "  "},
            {"session_length_minutes": 0},
            {"capacity": 0},
            {"end_minute": 8 * 60},
        ):
            with self.assertRaises(ClinicalValidationError):
                self.repo.create_opportunity(opportunity_input(**overrides))

    def test_unknown_tag_id_rejected(self):
        opp = self.repo.create_opportunity(opportunity_input())
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.set_opportunity_tags(opp.id, ["missing"])
        self.assertEqual(caught.exception.code, "TAG_NOT_FOUND")


class ClinicalRouterTests(ClinicalTestCase):
    payload = dict(
        site_name="Synthetic Clinic", day_of_week="Tuesday", start_minute=480, end_minute=660,
        session_length_minutes=60,
    )

    def test_opportunity_crud_through_routes(self):
        tag = clinical.create_tag(schemas.ClinicalTagIn(name="kids"), self.db)
        self.assertTrue(tag.is_unused)
        created = clinical.create_opportunity(
            schemas.PlacementOpportunityIn(**self.payload, tag_ids=[tag.id]), self.db
        )
        self.assertEqual(created.capacity, 1)
        self.assertEqual([t.name for t in created.tags], ["KIDS"])
        self.assertFalse(clinical.list_tags(self.db)[0].is_unused)
        updated = clinical.update_opportunity(
            created.id, schemas.PlacementOpportunityUpdate(**{**self.payload, "capacity": 3}), self.db
        )
        self.assertEqual(updated.capacity, 3)
        self.assertEqual(len(updated.tags), 1)
        cleared = clinical.set_opportunity_tags(created.id, schemas.PlacementOpportunityTagsIn(tag_ids=[]), self.db)
        self.assertEqual(cleared.tags, [])
        self.assertEqual(len(clinical.list_tags(self.db)), 1)
        clinical.delete_opportunity(created.id, self.db)
        self.assertEqual(clinical.list_opportunities(self.db), [])

    def test_error_status_mapping(self):
        clinical.create_tag(schemas.ClinicalTagIn(name="a"), self.db)
        with self.assertRaises(HTTPException) as dup:
            clinical.create_tag(schemas.ClinicalTagIn(name="A"), self.db)
        self.assertEqual(dup.exception.status_code, 409)
        with self.assertRaises(HTTPException) as missing:
            clinical.delete_tag("nope", self.db)
        self.assertEqual(missing.exception.status_code, 404)
        with self.assertRaises(HTTPException) as invalid:
            clinical.create_opportunity(
                schemas.PlacementOpportunityIn(**{**self.payload, "capacity": 0}), self.db
            )
        self.assertEqual(invalid.exception.status_code, 422)

    def test_delete_tag_reports_student_count(self):
        tag = clinical.create_tag(schemas.ClinicalTagIn(name="m"), self.db)
        student = self.repo.create_student(student_id="S1", name="Pat Test", email="")
        self.repo.set_tag_state(student.id, tag.id, TagState.PREFERRED)
        self.db.commit()
        self.assertEqual(clinical.list_tags(self.db)[0].student_count, 1)
        self.assertEqual(clinical.delete_tag(tag.id, self.db).affected_student_count, 1)


if __name__ == "__main__":
    unittest.main()
