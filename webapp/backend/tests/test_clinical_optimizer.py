"""Synthetic-data tests for the Clinical placement optimizer, persistence, and API."""

import unittest

from fastapi import HTTPException

from app.clinical_models import PlacementType, TagState
from app.routers import clinical
from app.services.availability import AvailabilityWindow
from app.services.clinical_candidates import generate_candidates
from app.services.clinical_optimizer import optimize
from app.services.clinical_scheduling_models import PreservedPlacement, TagPreference, UnplacedReason
from app.services.clinical_repository import ClinicalValidationError
from tests.test_clinical_candidates import H, build, opp, student, win
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

ALL_DAY = [win(0, 24 * H)]


def run(students, opportunities, preserved=(), tags=None):
    data = build(students, opportunities, tags=tags)
    return optimize(data, generate_candidates(data), preserved)


def assigned(result):
    return {p.student_id: p.opportunity_id for p in result.placements}


class OptimizerTests(unittest.TestCase):
    def test_maximizes_number_of_placed_students(self):
        # A greedy pass would give o1 to "a" and strand "b", who can only use o1.
        b = student("b", [win(9 * H - 15, 10 * H + 15)])
        o1 = opp("o1", start=9 * H, end=10 * H)
        o2 = opp("o2", day="Wednesday", start=9 * H, end=10 * H)
        a = student("a", ALL_DAY + [win(0, 24 * H, day="Wednesday")])
        result = run([a, b], [o1, o2])
        self.assertEqual(result.statistics.placed_count, 2)
        self.assertEqual(assigned(result), {"a": "o2", "b": "o1"})

    def test_capacity_enforced_and_reason_reported(self):
        students = [student(f"s{i}", ALL_DAY) for i in range(3)]
        result = run(students, [opp(capacity=2)])
        self.assertEqual(result.statistics.placed_count, 2)
        self.assertEqual([(u.student_id, u.reason) for u in result.unplaced], [("s2", UnplacedReason.CAPACITY_LIMIT)])

    def test_no_feasible_candidate_reason(self):
        result = run([student("a", ALL_DAY), student("b", [])], [opp()])
        self.assertEqual({u.student_id: u.reason for u in result.unplaced}, {"b": UnplacedReason.NO_FEASIBLE_CANDIDATE})

    def test_each_student_placed_at_most_once(self):
        result = run([student("a", ALL_DAY)], [opp("o1", capacity=3), opp("o2", capacity=3)])
        self.assertEqual(len(result.placements), 1)

    def test_preferred_tags_beat_tentative_and_start(self):
        prefs = {"t": TagPreference.PREFERRED}
        s = student("a", [win(0, 24 * H, "Tentative")], prefs)
        early = opp("early", start=8 * H, end=10 * H)
        tagged = opp("tagged", start=15 * H, end=17 * H, tags=["t"])
        result = run([s], [early, tagged])
        self.assertEqual(assigned(result), {"a": "tagged"})
        self.assertEqual(result.statistics.total_preferred_tag_matches, 1)

    def test_less_tentative_beats_earlier_start(self):
        s = student("a", [win(0, 12 * H, "Tentative"), win(12 * H, 24 * H, "Available")])
        early = opp("early", start=8 * H, end=10 * H)
        late = opp("late", start=14 * H, end=16 * H)
        result = run([s], [early, late])
        self.assertEqual(assigned(result), {"a": "late"})
        self.assertEqual(result.placements[0].tentative_minutes, 0)

    def test_earlier_start_is_last_preference(self):
        s = student("a", ALL_DAY)
        result = run([s], [opp("late", start=14 * H, end=16 * H), opp("early", start=8 * H, end=10 * H)])
        self.assertEqual(assigned(result), {"a": "early"})
        self.assertEqual(result.placements[0].selected_start_minute, 8 * H)

    def test_global_priorities_trade_start_for_preferred_total(self):
        prefs = {"t": TagPreference.PREFERRED}
        a = student("a", ALL_DAY, prefs)
        b = student("b", ALL_DAY)
        tagged = opp("tagged", start=15 * H, end=17 * H, tags=["t"])
        plain = opp("plain", start=8 * H, end=10 * H)
        result = run([a, b], [tagged, plain])
        self.assertEqual(assigned(result), {"a": "tagged", "b": "plain"})

    def test_ties_are_deterministic_and_input_order_independent(self):
        students = [student(f"s{i}", ALL_DAY) for i in range(4)]
        opportunities = [opp(f"o{i}") for i in range(2)]
        first = run(students, opportunities)
        again = run(students, opportunities)
        shuffled = run(list(reversed(students)), list(reversed(opportunities)))
        self.assertEqual(first, again)
        self.assertEqual(first.placements, shuffled.placements)
        # Canonical-order rule: with identical value, the earliest-named students are placed.
        self.assertEqual(sorted(assigned(first)), ["s0", "s1"])

    def test_preserved_manual_placement_consumes_capacity_and_excludes_student(self):
        students = [student("a", ALL_DAY), student("b", ALL_DAY), student("c", ALL_DAY)]
        preserved = [PreservedPlacement("a", "o1"), PreservedPlacement("c", None)]
        result = run(students, [opp(capacity=2)], preserved)
        self.assertEqual(assigned(result), {"b": "o1"})
        self.assertEqual(result.statistics.preserved_student_count, 2)

    def test_stamps_source_revision_and_rejects_mismatch(self):
        data = build([student("a", ALL_DAY)], [opp()], revision=7)
        self.assertEqual(optimize(data, generate_candidates(data)).source_data_revision, 7)
        other = build([student("a", ALL_DAY)], [opp()], revision=8)
        with self.assertRaises(ValueError):
            optimize(data, generate_candidates(other))

    def test_never_places_on_infeasible_candidates(self):
        tags = {"t": "T", "r": "R"}
        students = [
            student("ok", ALL_DAY),
            student("restricted", ALL_DAY, {"r": TagPreference.RESTRICTED}),
            student("unavailable", [win(0, 24 * H, "Unavailable")]),
            student("short", [win(9 * H, 9 * H + 30)]),
        ]
        opportunities = [opp("o1", tags=["r"], capacity=4), opp("o2", capacity=4)]
        data = build(students, opportunities, tags=tags)
        candidates = generate_candidates(data)
        result = optimize(data, candidates)
        feasible = {(c.student_id, c.opportunity_id): c for c in candidates.feasible}
        for placement in result.placements:
            candidate = feasible[(placement.student_id, placement.opportunity_id)]
            self.assertIn(placement.selected_start_minute, [f.start_minute for f in candidate.feasible_starts])
        self.assertEqual(assigned(result)["restricted"], "o2")
        self.assertIn("ok", assigned(result))
        self.assertNotIn("unavailable", assigned(result))
        self.assertNotIn("short", assigned(result))


class SchedulePersistenceTests(ClinicalTestCase):
    def setUp(self):
        super().setUp()
        self.site = self.repo.create_opportunity(opportunity_input(capacity=1))
        self.a = self.repo.create_student("A", "Alex", require_name=True)
        self.b = self.repo.create_student("B", "Blake", require_name=True)
        for item in (self.a, self.b):
            self.repo.set_student_availability(item.id, [AvailabilityWindow("Tuesday", 0, 24 * H, "Available")])
        self.db.commit()

    def test_generation_persists_automatic_placements_and_bumps_revision(self):
        result = self.repo.generate_schedule()
        self.db.commit()
        self.assertEqual(result.statistics.placed_count, 1)
        records = self.repo.list_placement_records()
        self.assertEqual([r.kind.value for r in records], ["automatic"])
        state = self.repo.get_state()
        self.assertEqual(state.schedule_revision, 1)
        self.assertFalse(state.is_schedule_stale)
        self.repo.generate_schedule()
        self.assertEqual(self.repo.get_state().schedule_revision, 2)
        self.assertEqual(len(self.repo.list_placement_records()), 1)

    def test_source_change_marks_schedule_stale_and_keeps_placements(self):
        self.repo.generate_schedule()
        self.db.commit()
        self.repo.create_student("C", "Casey", require_name=True)
        state = self.repo.get_state()
        self.assertTrue(state.is_schedule_stale)
        self.assertEqual(len(self.repo.list_placement_records()), 1)
        self.repo.generate_schedule()
        self.assertFalse(self.repo.get_state().is_schedule_stale)

    def test_regeneration_preserves_manual_decisions(self):
        self.repo.set_placement(self.a.id, PlacementType.MANUALLY_UNPLACED)
        self.repo.generate_schedule()
        kinds = {r.student_id: r.kind.value for r in self.repo.list_placement_records()}
        self.assertEqual(kinds, {self.a.id: "manuallyUnplaced", self.b.id: "automatic"})
        self.repo.set_placement(self.a.id, PlacementType.MANUAL, self.site.id, 9 * H)
        self.repo.generate_schedule()
        kinds = {r.student_id: r.kind.value for r in self.repo.list_placement_records()}
        self.assertEqual(kinds, {self.a.id: "manual"})  # manual placement filled the only seat

    def test_results_follow_source_data_changes(self):
        self.repo.generate_schedule()
        first = self.repo.list_placement_records()[0].student_id
        self.repo.set_student_availability(first, [AvailabilityWindow("Tuesday", 0, H, "Available")])
        self.repo.generate_schedule()
        self.assertNotEqual(self.repo.list_placement_records()[0].student_id, first)

    def test_clear_resets_revision_and_state(self):
        self.repo.generate_schedule()
        self.repo.create_student("C", "Casey", require_name=True)
        self.repo.clear_all_placements()
        state = self.repo.get_state()
        self.assertEqual((state.schedule_revision, state.is_schedule_stale), (0, False))
        self.assertEqual(self.repo.list_placement_records(), [])

    def test_generation_blocked_when_not_ready(self):
        self.repo.delete_student(self.a.id)
        self.repo.delete_student(self.b.id)
        self.db.commit()
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.generate_schedule()
        self.assertEqual(caught.exception.code, "NOT_READY")
        self.assertEqual(self.repo.get_state().schedule_revision, 0)


class ScheduleApiTests(ClinicalTestCase):
    def setUp(self):
        super().setUp()
        self.repo.create_opportunity(opportunity_input(capacity=2))
        for sid, name in (("A", "Alex"), ("B", "Blake"), ("C", "Casey")):
            item = self.repo.create_student(sid, name, require_name=True)
            if sid != "C":
                self.repo.set_student_availability(item.id, [AvailabilityWindow("Tuesday", 0, 24 * H, "Available")])
        self.db.commit()

    def test_generate_blocked_returns_structured_error(self):
        self.repo.clear_all_placements()
        for item in self.repo.list_students():
            self.repo.delete_student(item.id)
        self.db.commit()
        with self.assertRaises(HTTPException) as caught:
            clinical.generate_schedule(self.db)
        self.assertEqual(caught.exception.status_code, 409)
        self.assertEqual(caught.exception.detail["code"], "NOT_READY")

    def test_generate_get_and_clear(self):
        empty = clinical.get_schedule(self.db)
        self.assertFalse(empty.has_placements)
        out = clinical.generate_schedule(self.db)
        self.assertEqual((out.placed_count, out.unplaced_count), (2, 1))
        self.assertEqual(out.unplaced[0].student_name, "Casey")
        self.assertEqual(out.unplaced[0].reason, "NO_FEASIBLE_CANDIDATE")
        self.assertEqual(out.utilization, 1.0)
        self.assertEqual(clinical.get_schedule(self.db).schedule_revision, 1)
        cleared = clinical.clear_schedule(self.db)
        self.assertFalse(cleared.has_placements)
        self.assertEqual(cleared.schedule_revision, 0)


if __name__ == "__main__":
    unittest.main()
