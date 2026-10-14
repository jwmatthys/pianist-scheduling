"""Synthetic-data tests for Clinical Sites import parsing, validation, and replacement."""

import asyncio
import io
import json
import unittest
from datetime import time

import pandas as pd
from fastapi import HTTPException, UploadFile

from app.clinical_models import ClinicalStudent, PlacementType, TagState
from app.routers import clinical
from app.services import clinical_sites_importer as importer
from app.services.time_parsing import parse_time_of_day
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

HEADERS = ["Site", "Address", "Day", "Start", "End", "Session", "Travel", "Cap", "Pop", "Focus"]
MAPPING = {
    "site_name": "Site", "site_address": "Address", "day_of_week": "Day", "start_time": "Start",
    "end_time": "End", "session_length": "Session", "travel_time": "Travel", "capacity": "Cap",
}


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


GOOD_ROW = ["Synthetic Clinic", "1 Example Way", "tue", "8:00 AM", "5:00 PM", 60, 15, 2, "mental health", "Pediatric"]


def mapping(**overrides):
    return importer.ClinicalSitesImportMapping(**{**MAPPING, **overrides})


def parse(rows, **overrides):
    return importer.parse_sites("sites.csv", csv_bytes(rows), None, mapping(**overrides))


class ParsingTests(unittest.TestCase):
    def test_csv_fields_and_tags(self):
        site = parse([GOOD_ROW]).rows[0]
        self.assertEqual(site.errors, [])
        self.assertEqual(
            (site.site_name, site.day_of_week, site.start_minute, site.end_minute),
            ("Synthetic Clinic", "Tuesday", 480, 1020),
        )
        self.assertEqual((site.session_length_minutes, site.travel_minutes, site.capacity), (60, 15, 2))
        self.assertEqual(site.tags, ["MENTAL HEALTH", "PEDIATRIC"])

    def test_xlsx_with_excel_time_cells(self):
        buffer = io.BytesIO()
        frame = pd.DataFrame(
            [["Clinic", "Monday", time(9, 30), time(13, 0), 45]],
            columns=["Site", "Day", "Start", "End", "Session"],
        )
        frame.to_excel(buffer, index=False)
        parsed = importer.parse_sites(
            "sites.xlsx", buffer.getvalue(), None,
            importer.ClinicalSitesImportMapping(
                site_name="Site", day_of_week="Day", start_time="Start", end_time="End", session_length="Session"
            ),
        )
        site = parsed.rows[0]
        self.assertEqual((site.start_minute, site.end_minute, site.errors), (570, 780, []))
        self.assertEqual((site.capacity, site.travel_minutes), (1, 0))

    def test_xlsx_roundtrip_matches_csv(self):
        from_xlsx = importer.parse_sites("s.xlsx", xlsx_bytes([GOOD_ROW]), None, mapping()).rows[0]
        self.assertEqual(from_xlsx.tags, ["MENTAL HEALTH", "PEDIATRIC"])
        self.assertEqual(from_xlsx.start_minute, 480)

    def test_time_formats(self):
        cases = {"8:05": 485, "08:05": 485, "8:05 AM": 485, "8:05am": 485, "12:30 PM": 750, "1 PM": 780,
                 "13:00:00": 780, "12:00 AM": 0}
        for text, minutes in cases.items():
            self.assertEqual(parse_time_of_day(text), minutes, text)
        for bad in ("noon", "25:00", "13:00 PM", "8:75"):
            self.assertIsNone(parse_time_of_day(bad), bad)
        self.assertEqual(parse_time_of_day(0.5), 720)

    def test_capacity_defaults_when_unmapped_or_blank(self):
        unmapped = parse([GOOD_ROW], capacity=None).rows[0]
        self.assertEqual(unmapped.capacity, 1)
        blank = parse([GOOD_ROW[:7] + [None] + GOOD_ROW[8:]]).rows[0]
        self.assertEqual((blank.capacity, blank.errors), (1, []))

    def test_tags_one_cell_one_tag_no_splitting_and_collapse(self):
        row = GOOD_ROW[:8] + ["a, b; c", "  A, B; C  "]
        site = parse([row]).rows[0]
        self.assertEqual(site.tags, ["A, B; C"])
        row = GOOD_ROW[:8] + ["kids", "KIDS"]
        self.assertEqual(parse([row]).rows[0].tags, ["KIDS"])
        row = GOOD_ROW[:8] + ["", None]
        self.assertEqual(parse([row]).rows[0].tags, [])

    def test_unmapped_columns_become_tag_columns_regardless_of_heading(self):
        parsed = parse([GOOD_ROW])
        self.assertEqual(parsed.tag_columns, ["Pop", "Focus"])
        parsed = parse([GOOD_ROW], capacity=None)
        self.assertEqual(parsed.tag_columns, ["Cap", "Pop", "Focus"])
        self.assertEqual(parsed.rows[0].tags, ["2", "MENTAL HEALTH", "PEDIATRIC"])

    def test_blank_rows_skipped_and_row_numbers_match_sheet(self):
        parsed = parse([[None] * 10, GOOD_ROW])
        self.assertEqual([site.row_number for site in parsed.rows], [3])

    def test_inspect_suggests_mapping(self):
        headers = ["Site Name", "Day of Week", "Start Time", "End Time", "Session Length", "Misc"]
        inspection = importer.inspect_file("s.csv", csv_bytes([["A", "Mon", "8:00", "9:00", 30, "x"]], headers), None)
        self.assertEqual(inspection.suggested_mapping["site_name"], "Site Name")
        self.assertIsNone(inspection.suggested_mapping["capacity"])
        self.assertEqual(inspection.columns[-1], "Misc")


class ValidationTests(unittest.TestCase):
    def errors(self, **changes):
        row = list(GOOD_ROW)
        for index, value in changes.items():
            row[int(index)] = value
        return parse([row]).rows[0].errors

    def test_missing_required_fields(self):
        for column, label in ((0, "Site Name"), (2, "Day of Week"), (3, "Start Time"), (4, "End Time"), (5, "Session Length")):
            messages = self.errors(**{str(column): None})
            self.assertTrue(any(label in message and "required" in message for message in messages), label)

    def test_invalid_day_time_length_capacity(self):
        self.assertTrue(any("weekday" in m for m in self.errors(**{"2": "Funday"})))
        self.assertTrue(any("Start Time" in m and "recognized" in m for m in self.errors(**{"3": "soon"})))
        self.assertIn("End Time must be after Start Time.", self.errors(**{"4": "7:00 AM"}))
        self.assertIn("Session Length must be greater than 0.", self.errors(**{"5": 0}))
        self.assertIn("Session Length must be greater than 0.", self.errors(**{"5": -5}))
        self.assertTrue(any("whole number" in m for m in self.errors(**{"5": "abc"})))
        self.assertIn("Capacity must be greater than 0.", self.errors(**{"7": 0}))

    def test_invalid_values_are_not_silently_corrected(self):
        site = parse([list(GOOD_ROW[:5]) + [0] + GOOD_ROW[6:]]).rows[0]
        self.assertEqual(site.session_length_minutes, 0)
        preview = importer.build_preview(parse([GOOD_ROW[:5] + [0] + GOOD_ROW[6:]]), set(), 0, 0)
        self.assertFalse(preview.can_apply)

    def test_mapping_validation(self):
        data = csv_bytes([GOOD_ROW])
        with self.assertRaises(importer.ClinicalSitesImportError) as caught:
            importer.parse_sites("s.csv", data, None, mapping(site_name=None))
        self.assertEqual(caught.exception.code, "INVALID_MAPPING")
        with self.assertRaises(importer.ClinicalSitesImportError):
            importer.parse_sites("s.csv", data, None, mapping(site_name="Nope"))
        with self.assertRaises(importer.ClinicalSitesImportError):
            importer.parse_sites("s.csv", data, None, mapping(site_address="Site"))

    def test_preview_counts_and_tag_classification(self):
        bad = list(GOOD_ROW)
        bad[2] = "nope"
        preview = importer.build_preview(parse([GOOD_ROW, bad]), {"pediatric"}, 3, 4)
        self.assertEqual((preview.total_rows, preview.valid_rows, preview.invalid_rows), (2, 1, 1))
        self.assertEqual(preview.new_tags, ["MENTAL HEALTH"])
        self.assertEqual(preview.reused_tags, ["PEDIATRIC"])
        self.assertEqual((preview.existing_opportunity_count, preview.existing_placement_count), (3, 4))
        self.assertFalse(preview.can_apply)


def upload(data: bytes, name="sites.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


class ReplacementTests(ClinicalTestCase):
    def seed(self):
        keep = self.repo.create_tag("pediatric")
        self.repo.create_tag("orphan")
        old = self.repo.create_opportunity(opportunity_input(site_name="Old Site"), tag_ids=[keep.id])
        student = self.repo.create_student(student_id="S1", name="Pat Test", email="")
        self.repo.set_tag_state(student.id, keep.id, TagState.REQUIRED)
        self.repo.set_placement(student.id, PlacementType.MANUAL, old.id, 480)
        self.db.commit()
        return student

    def apply(self, rows, confirmed=True, **kwargs):
        return asyncio.run(
            clinical.apply_sites_import(
                file=upload(csv_bytes(rows)), mapping=json.dumps(MAPPING), sheet_name=None,
                confirmed=confirmed, db=self.db, **kwargs,
            )
        )

    def test_full_replacement_preserves_registry_and_students(self):
        student = self.seed()
        before = self.repo.get_state().source_data_revision
        result = self.apply([GOOD_ROW, ["Second", "", "Wed", "9:00", "11:00", 30, 0, 1, "", "PEDIATRIC"]])
        self.assertEqual(
            (result.opportunities_removed, result.placements_removed, result.opportunities_created, result.tags_created),
            (1, 1, 2, 1),
        )
        names = [o.site_name for o in self.repo.list_opportunities()]
        self.assertEqual(names, ["Second", "Synthetic Clinic"])
        tag_names = {t.name for t in self.repo.list_tags()}
        self.assertEqual(tag_names, {"PEDIATRIC", "ORPHAN", "MENTAL HEALTH"})
        summaries = {s.tag.name: s for s in self.repo.tag_summaries()}
        self.assertTrue(summaries["ORPHAN"].is_unused)
        self.assertEqual(summaries["PEDIATRIC"].opportunity_count, 2)
        self.assertEqual(self.repo.get_tag_state(student.id, self.repo.find_tag("pediatric").id), TagState.REQUIRED)
        self.assertIsNone(self.repo.get_placement(student.id))
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertEqual(state.schedule_revision, 0)
        self.assertFalse(state.is_schedule_stale)

    def test_tag_reuse_does_not_duplicate(self):
        self.seed()
        self.apply([["A", "", "Mon", "8:00", "9:00", 30, 0, 1, "PEDIATRIC", "pediatric"]])
        self.assertEqual(len([t for t in self.repo.list_tags() if t.name == "PEDIATRIC"]), 1)
        self.assertEqual([t.name for t in self.repo.list_opportunities()[0].tags], ["PEDIATRIC"])

    def test_requires_explicit_confirmation(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.apply([GOOD_ROW], confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual([o.site_name for o in self.repo.list_opportunities()], ["Old Site"])

    def test_invalid_rows_block_apply_and_change_nothing(self):
        self.seed()
        bad = list(GOOD_ROW)
        bad[5] = 0
        with self.assertRaises(HTTPException) as caught:
            self.apply([GOOD_ROW, bad])
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertTrue(any("Row 3" in issue for issue in caught.exception.detail["issues"]))
        self.assertEqual([o.site_name for o in self.repo.list_opportunities()], ["Old Site"])
        self.assertIsNotNone(self.repo.get_placement(self.db.query(ClinicalStudent).one().id))

    def test_preview_writes_nothing(self):
        self.seed()
        preview = asyncio.run(
            clinical.preview_sites_import(
                file=upload(csv_bytes([GOOD_ROW])), mapping=json.dumps(MAPPING), sheet_name=None, db=self.db
            )
        )
        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.existing_placement_count, 1)
        self.assertEqual(preview.new_tags, ["MENTAL HEALTH"])
        self.assertEqual(preview.reused_tags, ["PEDIATRIC"])
        self.db.rollback()
        self.assertEqual([o.site_name for o in self.repo.list_opportunities()], ["Old Site"])
        self.assertIsNone(self.repo.find_tag("mental health"))


if __name__ == "__main__":
    unittest.main()
