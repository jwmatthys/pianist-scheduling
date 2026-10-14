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

    # Node layout: 0 source, 1..S students, then opportunities, then sink.
    student_node = {s.id: 1 + index for index, s in enumerate(students)}
    opportunity_node = {o.id: 1 + len(students) + index for index, o in enumerate(opportunities)}
    source, sink = 0, 1 + len(students) + len(opportunities)
    flow = _Flow(sink + 1)
    for student in students:
        flow.add_edge(source, student_node[student.id], 1, _ZERO)
    pair_edges: list[tuple[int, Candidate]] = []
    position = 0
    for student in students:
        for opportunity in opportunities:
            candidate = feasible.get((student.id, opportunity.id))
            if candidate is None or remaining[opportunity.id] <= 0:
                continue
            assert candidate.tentative_minutes is not None and candidate.opportunity_start_minute is not None
            cost: Cost = (
                -candidate.preferred_tag_match_count,
                candidate.tentative_minutes,
                candidate.opportunity_start_minute,
                1 << position,
            )
            position += 1
            pair_edges.append((flow.add_edge(student_node[student.id], opportunity_node[opportunity.id], 1, cost), candidate))
    for opportunity in opportunities:
        flow.add_edge(opportunity_node[opportunity.id], sink, remaining[opportunity.id], _ZERO)
    flow.run(source, sink)

    placements = tuple(
        ProposedPlacement(
            student_id=candidate.student_id,
            opportunity_id=candidate.opportunity_id,
            selected_start_minute=candidate.opportunity_start_minute,  # type: ignore[arg-type]
            preferred_tag_match_count=candidate.preferred_tag_match_count,
            tentative_minutes=candidate.tentative_minutes,  # type: ignore[arg-type]
        )
        for edge, candidate in pair_edges
        if flow.cap[edge] == 0
    )
    placed_ids = {item.student_id for item in placements}
    has_feasible = {student_id for student_id, _ in feasible}
    unplaced = tuple(
        UnplacedStudent(
            s.id,
            UnplacedReason.CAPACITY_LIMIT if s.id in has_feasible else UnplacedReason.NO_FEASIBLE_CANDIDATE,
        )
        for s in students
        if s.id not in placed_ids
    )
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
