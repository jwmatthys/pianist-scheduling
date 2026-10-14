"""Synthetic-data tests for Clinical candidate generation and readiness analysis."""

import unittest

from app.clinical_models import PlacementType, TagState
from app.routers import clinical
from app.services.availability import AvailabilityWindow
from app.services.clinical_candidates import generate_candidates
from app.services.clinical_readiness import analyze_readiness
from app.services.clinical_scheduling_models import (
    CandidateIssueCode as Issue,
    OpportunityInput,
    ReadinessCode as Code,
    ReadinessSeverity,
    SchedulingInput,
    StudentInput,
    TagPreference,
)
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

H = 60


def opp(id="o1", day="Tuesday", start=9 * H, end=12 * H, session=60, travel=15, capacity=1, tags=()):
    return OpportunityInput(id, id, day, start, end, session, travel, capacity, frozenset(tags))


def student(id="s1", windows=(), prefs=None):
    return StudentInput(id, id, tuple(windows), prefs or {})


def win(start, end, status="Available", day="Tuesday"):
    return AvailabilityWindow(day, start, end, status)


def build(students, opportunities, tags=None, revision=1):
    return SchedulingInput(revision, tuple(students), tuple(opportunities), tags or {})


def one(students, opportunities, **kwargs):
    return generate_candidates(build(students, opportunities, **kwargs)).candidates[0]


class AvailabilityMatchingTests(unittest.TestCase):
    def test_available_window_yields_15_minute_starts_with_travel(self):
        # Commitment is 90 minutes; the window 8:45-12:15 allows starts 9:00..11:00.
        candidate = one([student(windows=[win(8 * H + 45, 12 * H + 15)])], [opp()])
        self.assertTrue(candidate.is_feasible)
        self.assertEqual(candidate.feasible_starts[0].start_minute, 9 * H)
        self.assertEqual(candidate.feasible_starts[-1].start_minute, 11 * H)
        self.assertEqual([s.start_minute % 15 for s in candidate.feasible_starts], [0] * len(candidate.feasible_starts))
        self.assertEqual(candidate.tentative_minutes, 0)
        self.assertEqual(candidate.opportunity_start_minute, 9 * H)

    def test_exact_boundary_fits_and_one_minute_short_does_not(self):
        exact = [win(9 * H - 15, 10 * H + 15)]
        self.assertTrue(one([student(windows=exact)], [opp()]).is_feasible)
        short = [win(9 * H - 15, 10 * H + 14)]
        self.assertFalse(one([student(windows=short)], [opp()]).is_feasible)
        early = [win(9 * H - 14, 10 * H + 15)]
        self.assertFalse(one([student(windows=early)], [opp()]).is_feasible)

    def test_travel_is_required_before_and_after(self):
        no_travel_before = [win(9 * H, 10 * H + 15)]
        no_travel_after = [win(8 * H + 45, 10 * H)]
        self.assertFalse(one([student(windows=no_travel_before)], [opp(end=10 * H)]).is_feasible)
        self.assertFalse(one([student(windows=no_travel_after)], [opp(end=10 * H)]).is_feasible)
        self.assertTrue(one([student(windows=no_travel_before)], [opp(end=10 * H, travel=0)]).is_feasible)

    def test_session_must_fit_inside_site_window(self):
        # Student is free all day, but the site window is exactly one session long.
        candidate = one([student(windows=[win(0, 24 * H)])], [opp(start=9 * H, end=10 * H)])
        self.assertEqual([s.start_minute for s in candidate.feasible_starts], [9 * H])

    def test_wrong_day_is_not_feasible(self):
        candidate = one([student(windows=[win(0, 24 * H, day="Monday")])], [opp()])
        self.assertFalse(candidate.is_feasible)
        self.assertIn(Issue.NO_AVAILABILITY_FIT, candidate.issue_codes)

    def test_tentative_is_feasible_and_counts_minutes(self):
        windows = [win(8 * H + 45, 9 * H + 45), win(9 * H + 45, 10 * H + 15, "Tentative")]
        candidate = one([student(windows=windows)], [opp(end=10 * H)])
        self.assertTrue(candidate.is_feasible)
        self.assertEqual(candidate.tentative_minutes, 30)

    def test_mixed_available_and_tentative_counts_only_tentative(self):
        windows = [win(8 * H + 45, 9 * H + 30, "Tentative"), win(9 * H + 30, 10 * H + 15)]
        candidate = one([student(windows=windows)], [opp(end=10 * H)])
        self.assertEqual(candidate.tentative_minutes, 45)

    def test_best_start_minimizes_tentative_before_earliest(self):
        windows = [win(8 * H + 45, 10 * H, "Tentative"), win(10 * H, 12 * H + 15)]
        candidate = one([student(windows=windows)], [opp()])
        self.assertEqual(candidate.tentative_minutes, 0)
        self.assertEqual(candidate.opportunity_start_minute, 10 * H + 15)
        self.assertGreater(len(candidate.feasible_starts), 1)
        self.assertEqual(candidate.feasible_starts[0].start_minute, 9 * H)
        self.assertGreater(candidate.feasible_starts[0].tentative_minutes, 0)

    def test_unavailable_prohibits_placement(self):
        windows = [win(8 * H, 12 * H), win(9 * H + 30, 9 * H + 45, "Unavailable")]
        candidate = one([student(windows=windows)], [opp(start=9 * H, end=10 * H + 15, session=60)])
        # Starts at 9:00 and 9:15 are blocked by the Unavailable minutes in the commitment.
        self.assertFalse(candidate.is_feasible)

    def test_unavailable_overrides_overlapping_available(self):
        windows = [win(8 * H, 12 * H), win(9 * H, 10 * H, "Unavailable")]
        self.assertFalse(one([student(windows=windows)], [opp(end=10 * H)]).is_feasible)

    def test_missing_availability_is_not_available(self):
        candidate = one([student(windows=[])], [opp()])
        self.assertFalse(candidate.is_feasible)
        self.assertIn(Issue.STUDENT_WITHOUT_AVAILABILITY, candidate.issue_codes)

    def test_travel_cannot_cross_midnight(self):
        candidate = one([student(windows=[win(0, 24 * H)])], [opp(start=0, end=H, session=60)])
        self.assertFalse(candidate.is_feasible)

    def test_non_30_minute_windows_are_supported(self):
        candidate = one([student(windows=[win(8 * H + 55, 10 * H + 5)])], [opp(end=10 * H, travel=5)])
        self.assertEqual([s.start_minute for s in candidate.feasible_starts], [9 * H])


class TagRuleTests(unittest.TestCase):
    ALL_DAY = [win(0, 24 * H)]

    def test_required_tags_use_and_semantics(self):
        prefs = {"a": TagPreference.REQUIRED, "b": TagPreference.REQUIRED}
        self.assertFalse(one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["a"])]).is_feasible)
        both = one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["a", "b", "c"])])
        self.assertTrue(both.is_feasible)
        missing = one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["a"])])
        issue = next(i for i in missing.issues if i.code is Issue.REQUIRED_TAG_MISSING)
        self.assertEqual(issue.tag_ids, ("b",))

    def test_any_restricted_tag_disallows(self):
        prefs = {"r": TagPreference.RESTRICTED}
        blocked = one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["x", "r"])])
        self.assertFalse(blocked.is_feasible)
        self.assertIn(Issue.RESTRICTED_TAG_PRESENT, blocked.issue_codes)
        self.assertTrue(one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["x"])]).is_feasible)

    def test_preferred_tags_count_but_do_not_affect_feasibility(self):
        prefs = {"p1": TagPreference.PREFERRED, "p2": TagPreference.PREFERRED}
        two = one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=["p1", "p2"])])
        none = one([student(windows=self.ALL_DAY, prefs=prefs)], [opp(tags=[])])
        self.assertEqual((two.is_feasible, two.preferred_tag_match_count), (True, 2))
        self.assertEqual((none.is_feasible, none.preferred_tag_match_count), (True, 0))

    def test_neutral_tags_are_ignored(self):
        self.assertTrue(one([student(windows=self.ALL_DAY)], [opp(tags=["anything"])]).is_feasible)

    def test_all_reasons_are_reported_together(self):
        prefs = {"need": TagPreference.REQUIRED}
        candidate = one([student(windows=[win(0, H)], prefs=prefs)], [opp()])
        self.assertEqual(candidate.issue_codes, {Issue.REQUIRED_TAG_MISSING, Issue.NO_AVAILABILITY_FIT})


class CapacityAndStructureTests(unittest.TestCase):
    def test_capacity_is_exposed_but_not_enforced(self):
        students = [student(f"s{i}", windows=[win(0, 24 * H)]) for i in range(3)]
        result = generate_candidates(build(students, [opp(capacity=1)]))
        self.assertEqual(len(result.feasible), 3)
        self.assertEqual({c.opportunity_capacity for c in result.candidates}, {1})

    def test_invalid_site_windows_and_zero_capacity_have_no_candidates(self):
        s = student(windows=[win(0, 24 * H)])
        for bad, issue in [
            (opp(start=10 * H, end=9 * H), Issue.INVALID_SITE_WINDOW),
            (opp(session=0), Issue.INVALID_SITE_WINDOW),
            (opp(start=9 * H, end=9 * H + 30, session=60), Issue.INVALID_SITE_WINDOW),
            (opp(capacity=0), Issue.ZERO_CAPACITY),
        ]:
            candidate = one([s], [bad])
            self.assertFalse(candidate.is_feasible)
            self.assertEqual(candidate.issue_codes, {issue})

    def test_output_is_deterministic_and_ordered(self):
        students = [student(f"s{i}", windows=[win(0, 24 * H)]) for i in range(3)]
        opportunities = [opp("a"), opp("b")]
        first = generate_candidates(build(students, opportunities))
        self.assertEqual(first, generate_candidates(build(students, opportunities)))
        self.assertEqual([(c.student_id, c.opportunity_id) for c in first.candidates][:3], [("s0", "a"), ("s0", "b"), ("s1", "a")])

    def test_helpers_and_revision_stamp(self):
        result = generate_candidates(build([student("s1", windows=[win(0, 24 * H)]), student("s2")], [opp()], revision=7))
        self.assertEqual(result.source_data_revision, 7)
        self.assertEqual(len(result.feasible_for_student("s1")), 1)
        self.assertEqual(result.feasible_for_student("s2"), ())


def report(students, opportunities, tags=None):
    data = build(students, opportunities, tags)
    return analyze_readiness(data, generate_candidates(data))


def codes(r):
    return [w.code for w in r.warnings]


class ReadinessTests(unittest.TestCase):
    def test_empty_inputs_block(self):
        r = report([], [])
        self.assertEqual(set(codes(r)), {Code.NO_STUDENTS, Code.NO_PLACEMENT_OPPORTUNITIES})
        self.assertFalse(r.is_ready)

    def test_clean_input_is_ready_without_warnings(self):
        r = report([student(windows=[win(0, 24 * H)])], [opp()])
        self.assertEqual(r.warnings, ())
        self.assertTrue(r.is_ready)
        self.assertEqual((r.student_count, r.opportunity_count, r.feasible_candidate_count), (1, 1, 1))

    def test_student_without_availability_is_separate_from_without_candidates(self):
        students = [
            student("none"),
            student("wrongday", windows=[win(0, 24 * H, day="Friday")]),
            student("ok", windows=[win(0, 24 * H)]),
        ]
        r = report(students, [opp()])
        by_code = {c: [w.subject.id for w in r.warnings if w.code is c and w.subject] for c in set(codes(r))}
        self.assertEqual(by_code[Code.STUDENT_WITHOUT_AVAILABILITY], ["none"])
        self.assertEqual(by_code[Code.STUDENT_WITHOUT_CANDIDATES], ["wrongday"])
        self.assertTrue(r.is_ready)

    def test_student_whose_tags_exclude_everything_has_no_candidates(self):
        s = student("s", windows=[win(0, 24 * H)], prefs={"r": TagPreference.RESTRICTED})
        r = report([s], [opp(tags=["r"])], {"r": "R"})
        self.assertIn(Code.STUDENT_WITHOUT_CANDIDATES, codes(r))
        self.assertIn(Code.NO_FEASIBLE_CANDIDATES, codes(r))
        self.assertFalse(r.is_ready)

    def test_opportunity_without_availability_match_ignores_tags(self):
        students = [student("s", windows=[win(0, 24 * H)], prefs={"x": TagPreference.REQUIRED})]
        r = report(students, [opp("fit"), opp("friday", day="Friday")], {"x": "X"})
        flagged = [w.subject.id for w in r.warnings if w.code is Code.OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH]
        self.assertEqual(flagged, ["friday"])

    def test_required_tag_without_matching_site(self):
        students = [student("s", windows=[win(0, 24 * H)], prefs={"mh": TagPreference.REQUIRED, "ok": TagPreference.REQUIRED})]
        r = report(students, [opp(tags=["ok"])], {"mh": "MENTAL HEALTH", "ok": "OK"})
        tagged = [(w.subject.id, w.subject.label) for w in r.warnings if w.code is Code.REQUIRED_TAG_WITHOUT_MATCH]
        self.assertEqual(tagged, [("mh", "MENTAL HEALTH")])

    def test_zero_capacity_and_invalid_window_are_errors(self):
        r = report([student(windows=[win(0, 24 * H)])], [opp("zero", capacity=0), opp("bad", start=10 * H, end=9 * H), opp("good")])
        severities = {w.code: w.severity for w in r.warnings}
        self.assertEqual(severities[Code.ZERO_CAPACITY], ReadinessSeverity.ERROR)
        self.assertEqual(severities[Code.INVALID_SITE_WINDOW], ReadinessSeverity.ERROR)
        self.assertNotIn(Code.OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH, severities)

    def test_capacity_below_student_count_is_informational(self):
        students = [student(f"s{i}", windows=[win(0, 24 * H)]) for i in range(3)]
        r = report(students, [opp(capacity=2)])
        info = [w for w in r.warnings if w.code is Code.CAPACITY_BELOW_STUDENT_COUNT]
        self.assertEqual(len(info), 1)
        self.assertEqual(info[0].severity, ReadinessSeverity.INFO)
        self.assertTrue(r.is_ready)


class RepositoryIntegrationTests(ClinicalTestCase):
    def setUp(self):
        super().setUp()
        self.site = self.repo.create_opportunity(opportunity_input(day_of_week="Tue", capacity=2))
        self.student = self.repo.create_student("S1", "Pat", require_name=True)

    def analyze(self):
        data = self.repo.scheduling_input()
        return data, generate_candidates(data)

    def test_snapshot_reflects_persisted_data_with_readable_labels(self):
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Tuesday", 7 * H, 12 * H, "Available")])
        data, candidates = self.analyze()
        self.assertEqual(data.opportunities[0].label, "Synthetic Clinic, Tue 8:00 AM-11:00 AM")
        self.assertEqual(data.opportunities[0].day, "Tuesday")
        self.assertEqual(len(candidates.feasible), 1)

    def test_tag_states_map_and_neutral_is_absent(self):
        required = self.repo.create_tag("kids")
        preferred = self.repo.create_tag("adults")
        self.repo.set_opportunity_tags(self.site.id, [preferred.id])
        self.repo.set_student_tag_states(self.student.id, {required.id: TagState.REQUIRED, preferred.id: TagState.PREFERRED})
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Tuesday", 0, 24 * H, "Available")])
        data, candidates = self.analyze()
        self.assertEqual(data.students[0].tag_preferences, {required.id: TagPreference.REQUIRED, preferred.id: TagPreference.PREFERRED})
        self.assertEqual(candidates.candidates[0].issue_codes, {Issue.REQUIRED_TAG_MISSING})
        self.assertEqual(candidates.candidates[0].preferred_tag_match_count, 0)

    def test_every_source_change_advances_the_revision_stamp(self):
        revisions = [self.repo.scheduling_input().source_data_revision]

        def changed(action):
            action()
            revisions.append(self.repo.scheduling_input().source_data_revision)

        tag = self.repo.create_tag("kids")
        revisions.append(self.repo.scheduling_input().source_data_revision)
        changed(lambda: self.repo.set_opportunity_tags(self.site.id, [tag.id]))
        changed(lambda: self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Tuesday", 0, H, "Available")]))
        changed(lambda: self.repo.set_student_tag_states(self.student.id, {tag.id: TagState.PREFERRED}))
        changed(lambda: self.repo.update_student(self.student.id, "S1", "Pat Two"))
        changed(lambda: self.repo.create_student("S2", "Sam", require_name=True))
        changed(lambda: self.repo.rename_tag(tag.id, "children"))
        changed(lambda: self.repo.delete_tag(tag.id))
        self.assertEqual(revisions, sorted(set(revisions)))

    def test_results_change_after_edits_and_do_not_use_placements(self):
        data, before = self.analyze()
        self.assertEqual(len(before.feasible), 0)
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Tuesday", 0, 24 * H, "Available")])
        data_after, after = self.analyze()
        self.assertGreater(data_after.source_data_revision, data.source_data_revision)
        self.assertEqual(len(after.feasible), 1)
        # Existing placements do not alter candidates: capacity is not enforced here.
        self.repo.set_placement(self.student.id, PlacementType.MANUAL, self.site.id, 9 * H)
        self.assertEqual(len(self.analyze()[1].feasible), 1)

    def test_readiness_endpoint_returns_structured_report(self):
        self.db.commit()
        out = clinical.get_readiness(self.db)
        self.assertEqual((out.student_count, out.opportunity_count, out.feasible_candidate_count), (1, 1, 0))
        self.assertFalse(out.is_ready)
        codes_out = {w.code for w in out.warnings}
        self.assertIn("STUDENT_WITHOUT_AVAILABILITY", codes_out)
        self.assertIn("NO_FEASIBLE_CANDIDATES", codes_out)
        warning = next(w for w in out.warnings if w.code == "STUDENT_WITHOUT_AVAILABILITY")
        self.assertEqual((warning.subject.kind, warning.subject.label), ("student", "Pat"))
        before = out.source_data_revision
        clinical.set_student_availability(
            self.student.id,
            clinical.schemas.ClinicalStudentAvailabilityIO(windows=[
                clinical.schemas.ClinicalAvailabilityWindowIO(day="Tuesday", start_minute=0, end_minute=1440, status="Available")
            ]),
            self.db,
        )
        out = clinical.get_readiness(self.db)
        self.assertTrue(out.is_ready)
        self.assertEqual(out.students_with_candidates, 1)
        self.assertGreater(out.source_data_revision, before)


if __name__ == "__main__":
    unittest.main()
