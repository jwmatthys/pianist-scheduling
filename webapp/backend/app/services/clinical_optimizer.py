"""Clinical placement optimizer.

Pure and database-free: it consumes a ``SchedulingInput`` and its ``CandidateSet``
and chooses among already-legal candidates. It never re-checks legality, so
feasibility, tag, availability, and travel rules are inherited from candidate
generation; capacity is enforced here.

Objective (spec 8.4), compared lexicographically and never blended into one score:

1. maximize the number of students placed;
2. maximize total satisfied Preferred student/tag pairs;
3. minimize total Tentative minutes (travel before + session + travel after);
4. minimize the sum of selected session start minutes;
5. resolve remaining ties with the canonical-order rule below.

Method: exact minimum-cost maximum flow (source -> student -> opportunity -> sink)
by successive shortest paths. Flow value is priority 1. Edge costs are tuples
``(-preferred, tentative, start, tie_rank)`` that are added and compared
component-wise and lexicographically, which is exact for priorities 2-5. Because the
objective is global, a flexible student cannot take the only legal opportunity of a
hard-to-place student (no roster-order greedy).

Selected start. Per student/opportunity pair the start is the candidate's own
``opportunity_start_minute``: the feasible start with the fewest Tentative minutes,
earliest on ties. Starts are independent across students (capacity is a total count,
not simultaneous occupancy), so this is optimal for priorities 3 and 4.

Tie-break rule (priority 5). Students are ordered by (name casefolded, id) and
opportunities by (label casefolded, id); feasible pairs are ordered by (student
position, opportunity position). Each pair carries a distinct weight 2**position,
so every distinct set of placements has a distinct total. The schedule with the
smaller total wins, i.e. compare the assigned pair sets from the last pair to the
first in that order and prefer the set that avoids the later pair. The rule uses no
clinical attribute and does not depend on input order.
"""

from __future__ import annotations

from collections import deque
from typing import Sequence

from .clinical_scheduling_models import (
    Candidate,
    CandidateSet,
    OptimizationResult,
    OptimizationStatistics,
    OpportunityInput,
    PreservedPlacement,
    ProposedPlacement,
    SchedulingInput,
    StudentInput,
    UnplacedReason,
    UnplacedStudent,
    commitment_interval,
)

Cost = tuple[int, int, int, int]
_ZERO: Cost = (0, 0, 0, 0)


def _add(a: Cost, b: Cost) -> Cost:
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2], a[3] + b[3])


def _neg(a: Cost) -> Cost:
    return (-a[0], -a[1], -a[2], -a[3])


class _Flow:
    """Min-cost max-flow with tuple costs (an ordered group, so SSP stays exact)."""

    def __init__(self, node_count: int):
        self.graph: list[list[int]] = [[] for _ in range(node_count)]
        self.to: list[int] = []
        self.cap: list[int] = []
        self.cost: list[Cost] = []

    def add_edge(self, source: int, target: int, capacity: int, cost: Cost) -> int:
        index = len(self.to)
        self.graph[source].append(index)
        self.to.append(target)
        self.cap.append(capacity)
        self.cost.append(cost)
        self.graph[target].append(index + 1)
        self.to.append(source)
        self.cap.append(0)
        self.cost.append(_neg(cost))
        return index

    def run(self, source: int, sink: int) -> None:
        node_count = len(self.graph)
        while True:
            dist: list[Cost | None] = [None] * node_count
            parent_edge = [-1] * node_count
            dist[source] = _ZERO
            queue = deque([source])
            queued = [False] * node_count
            queued[source] = True
            while queue:
                node = queue.popleft()
                queued[node] = False
                base = dist[node]
                for edge in self.graph[node]:
                    if self.cap[edge] <= 0:
                        continue
                    target = self.to[edge]
                    candidate = _add(base, self.cost[edge])
                    if dist[target] is None or candidate < dist[target]:
                        dist[target] = candidate
                        parent_edge[target] = edge
                        if not queued[target]:
                            queued[target] = True
                            queue.append(target)
            if dist[sink] is None:
                return
            node = sink
            while node != source:
                edge = parent_edge[node]
                self.cap[edge] -= 1
                self.cap[edge ^ 1] += 1
                node = self.to[edge ^ 1]


def _student_key(student: StudentInput) -> tuple[str, str]:
    return (student.label.casefold(), student.id)


def _opportunity_key(opportunity: OpportunityInput) -> tuple[str, str]:
    return (opportunity.label.casefold(), opportunity.id)


def optimize(
    scheduling_input: SchedulingInput,
    candidates: CandidateSet,
    preserved: Sequence[PreservedPlacement] = (),
) -> OptimizationResult:
    if candidates.source_data_revision != scheduling_input.source_data_revision:
        raise ValueError("The candidate set was generated from a different source-data revision.")

    held = {item.student_id for item in preserved}
    consumed: dict[str, int] = {}
    for item in preserved:
        if item.opportunity_id is not None:
            consumed[item.opportunity_id] = consumed.get(item.opportunity_id, 0) + 1

    students = sorted((s for s in scheduling_input.students if s.id not in held), key=_student_key)
    opportunities = sorted(scheduling_input.opportunities, key=_opportunity_key)
    remaining = {o.id: max(0, o.capacity - consumed.get(o.id, 0)) for o in opportunities}
    eligible_ids = {s.id for s in students}
    feasible: dict[tuple[str, str], Candidate] = {
        (c.student_id, c.opportunity_id): c
        for c in candidates.candidates
        if c.is_feasible and c.student_id in eligible_ids
    }

    opportunity_by_id = {o.id: o for o in opportunities}
    person_of = {s.id: s.person_key or s.id for s in scheduling_input.students}

    def interval(day: str, start: int, opportunity: OpportunityInput) -> tuple[str, int, int]:
        low, high = commitment_interval(opportunity, start)
        return (day, low, high)

    def overlaps(first: tuple[str, int, int], second: tuple[str, int, int]) -> bool:
        return first[0] == second[0] and first[1] < second[2] and second[1] < first[2]

    # Commitments (travel + session + travel) already held by a person's other records are fixed obstacles.
    held_intervals: dict[str, list[tuple[str, int, int]]] = {}
    for item in preserved:
        opportunity = opportunity_by_id.get(item.opportunity_id) if item.opportunity_id else None
        if opportunity is not None and item.start_minute is not None:
            person = person_of.get(item.student_id, item.student_id)
            held_intervals.setdefault(person, []).append(interval(opportunity.day, item.start_minute, opportunity))

    # Node layout: 0 source, 1..S students, then opportunities, then sink.
    student_node = {s.id: 1 + index for index, s in enumerate(students)}
    opportunity_node = {o.id: 1 + len(students) + index for index, o in enumerate(opportunities)}
    source, sink = 0, 1 + len(students) + len(opportunities)
    pairs: list[tuple[StudentInput, OpportunityInput, Candidate, Cost, tuple[str, int, int]]] = []
    for student in students:
        for opportunity in opportunities:
            candidate = feasible.get((student.id, opportunity.id))
            if candidate is None or remaining[opportunity.id] <= 0:
                continue
            assert candidate.tentative_minutes is not None and candidate.opportunity_start_minute is not None
            slot = interval(opportunity.day, candidate.opportunity_start_minute, opportunity)
            if any(overlaps(slot, held) for held in held_intervals.get(person_of[student.id], ())):
                continue
            cost: Cost = (
                -candidate.preferred_tag_match_count,
                candidate.tentative_minutes,
                candidate.opportunity_start_minute,
                1 << len(pairs),
            )
            pairs.append((student, opportunity, candidate, cost, slot))

    def solve(forbidden: frozenset[int]) -> tuple[list[int], tuple[int, Cost]]:
        flow = _Flow(sink + 1)
        for student in students:
            flow.add_edge(source, student_node[student.id], 1, _ZERO)
        edges: list[tuple[int, int]] = []
        for index, (student, opportunity, _, cost, _) in enumerate(pairs):
            if index not in forbidden:
                edges.append((flow.add_edge(student_node[student.id], opportunity_node[opportunity.id], 1, cost), index))
        for opportunity in opportunities:
            flow.add_edge(opportunity_node[opportunity.id], sink, remaining[opportunity.id], _ZERO)
        flow.run(source, sink)
        chosen = [index for edge, index in edges if flow.cap[edge] == 0]
        total = _ZERO
        for index in chosen:
            total = _add(total, pairs[index][3])
        return chosen, (-len(chosen), total)

    def first_person_conflict(chosen: list[int]) -> tuple[int, int] | None:
        for position, first in enumerate(chosen):
            for second in chosen[position + 1:]:
                if (
                    person_of[pairs[first][0].id] == person_of[pairs[second][0].id]
                    and overlaps(pairs[first][4], pairs[second][4])
                ):
                    return first, second
        return None

    # Records that share a person key must never hold overlapping commitments (travel included). That is not a flow
    # constraint, so solve the flow relaxation and branch on each violated pair; every feasible
    # schedule leaves out at least one of the two, so the search is exact. Without such
    # conflicts the first solve is final.
    best: list[tuple[tuple[int, Cost], list[int]] | None] = [None]

    def branch(forbidden: frozenset[int]) -> None:
        chosen, key = solve(forbidden)
        if best[0] is not None and key >= best[0][0]:
            return
        conflict = first_person_conflict(chosen)
        if conflict is None:
            best[0] = (key, chosen)
            return
        branch(forbidden | {conflict[0]})
        branch(forbidden | {conflict[1]})

    branch(frozenset())
    chosen_indices = best[0][1] if best[0] is not None else []

    placements = tuple(
        ProposedPlacement(
            student_id=candidate.student_id,
            opportunity_id=candidate.opportunity_id,
            selected_start_minute=candidate.opportunity_start_minute,  # type: ignore[arg-type]
            preferred_tag_match_count=candidate.preferred_tag_match_count,
            tentative_minutes=candidate.tentative_minutes,  # type: ignore[arg-type]
        )
        for index, (_, _, candidate, _, _) in enumerate(pairs)
        if index in set(chosen_indices)
    )
    placed_ids = {item.student_id for item in placements}
    has_feasible = {student_id for student_id, _ in feasible}
    commitments_by_person: dict[str, list[tuple[str, int, int]]] = {
        person: list(items) for person, items in held_intervals.items()
    }
    for index in chosen_indices:
        commitments_by_person.setdefault(person_of[pairs[index][0].id], []).append(pairs[index][4])

    def blocked_by_same_person(student_id: str) -> bool:
        """True when every feasible opportunity collides with another commitment of the same person."""
        taken = commitments_by_person.get(person_of[student_id], ())
        options = [
            interval(opportunity_by_id[opportunity_id].day, candidate.opportunity_start_minute, opportunity_by_id[opportunity_id])
            for (owner, opportunity_id), candidate in feasible.items()
            if owner == student_id
        ]
        return bool(options) and all(any(overlaps(option, other) for other in taken) for option in options)

    def reason_for(student_id: str) -> UnplacedReason:
        if student_id not in has_feasible:
            return UnplacedReason.NO_FEASIBLE_CANDIDATE
        if blocked_by_same_person(student_id):
            return UnplacedReason.STUDENT_OVERLAP
        return UnplacedReason.CAPACITY_LIMIT

    unplaced = tuple(UnplacedStudent(s.id, reason_for(s.id)) for s in students if s.id not in placed_ids)
    return OptimizationResult(
        source_data_revision=scheduling_input.source_data_revision,
        placements=placements,
        unplaced=unplaced,
        statistics=OptimizationStatistics(
            eligible_student_count=len(students),
            placed_count=len(placements),
            unplaced_count=len(unplaced),
            preserved_student_count=len(held),
            total_preferred_tag_matches=sum(item.preferred_tag_match_count for item in placements),
            total_tentative_minutes=sum(item.tentative_minutes for item in placements),
            total_start_minutes=sum(item.selected_start_minute for item in placements),
        ),
    )
