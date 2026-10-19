# Clinical Placement Scheduling Module Overview and Implementation Specification

**Status:** Approved module design for implementation.  
**Product:** Music Program Scheduler  
**Module:** Clinical Placement Scheduling / Clinical Placements

## 1. Purpose and implementation guardrails

The Clinical Placement Scheduling module places a roster of students, each with recurring weekly availability and placement requirements/preferences/restrictions, into Clinical Site Placement Opportunities.

This module is independent of Accompanist Scheduling and Performance Juries. It does not consume data or results from either module and does not publish data required by either module. It belongs to the active Scheduling Session and must follow the application's existing local-only privacy, session, migration, import, validation, and reporting architecture.

During initial Clinical implementation:

- Treat the existing Pianist/Accompanist Scheduling and Jury Scheduling modules as **feature-frozen for user testing**.
- Do not redesign, refactor, rename, or otherwise change their behavior except where a narrowly scoped extraction/generalization is required to reuse already-proven shared infrastructure.
- Any shared-component extraction must preserve existing behavior and pass the existing module tests.
- Do not create a universal optimizer, universal assignment type, universal Student record, or universal spreadsheet schema.
- Keep Clinical entities, constraints, optimizer, placement result, validation policy, and reports module-owned.
- All scheduling/imported data remains on the user's device. Do not add cloud services, telemetry, analytics, Graph/Forms APIs, or other network integrations.

## 2. Module navigation and workflow

The Clinical Placement module has four tabs in this order:

1. **Clinical Sites**
2. **Student Roster**
3. **Schedule**
4. **Reports**

Expected workflow:

1. Enter or import Placement Opportunities.
2. Review/edit Placement Opportunities and manage tags.
3. Enter or import the student roster and recurring weekly availability.
4. Review/edit each student's availability and tag requirements/preferences/restrictions.
5. Generate Placements.
6. Review unplaced students and warnings; make manual placements or time adjustments as needed.
7. Generate/review reports and export Markdown or PDF.

## 3. Clinical Sites and Placement Opportunities

### 3.1 Placement Opportunity is the scheduling unit

A spreadsheet row represents one independent **Placement Opportunity**. Rows must **not** be consolidated merely because they share a site name, address, or contact information. The same physical site may offer multiple windows with different times, capacities, or tags.

Use a flat list in the Clinical Sites UI, with one row per Placement Opportunity. Do not initially group rows by site.

### 3.2 Placement Opportunity fields

Each Placement Opportunity contains:

- internal application-generated identifier, never required from or displayed to the user;
- site name;
- site address;
- site contact name;
- site contact email;
- site contact phone;
- recurring day of week;
- opportunity start time;
- opportunity end time;
- session length;
- travel time;
- capacity;
- zero or more tags.

Capacity defaults to **1** when no capacity is supplied.

Travel time is symmetric: the same travel duration is required immediately before and immediately after the clinical session.

Capacity is the **total number of students who may be automatically assigned to that Placement Opportunity**, not a simultaneous-occupancy limit. For example, capacity 2 means at most two students total may be automatically assigned to that opportunity. Their clinical sessions may overlap completely or occur at different times within the opportunity window.

### 3.3 Manual Clinical Site editing

After import, Placement Opportunities are ordinary editable application data. The user may:

- add a Placement Opportunity manually;
- edit any Placement Opportunity field;
- edit capacity, session length, travel time, day, and times;
- add/remove tags;
- delete a Placement Opportunity.

Editing source data must not silently delete an existing schedule. See Section 9.

## 4. Clinical Site import

Accept local CSV and XLSX through the application's existing local import/mapping/preview infrastructure.
Tag behavior follows the canonical [Shared Tag Architecture](SHARED-TAG-ARCHITECTURE.md).
Each imported tag column must be explicitly mapped to the Clinical `Tags` field;
unmapped columns never become tags.

The intended source data includes the standard fields:

- Site name
- Site address
- Site contact name
- Site contact email
- Site contact phone
- Day of week
- Start time
- End time
- Session length
- Travel time
- Capacity (optional; default 1)
- Tags (optional; explicitly mapped)

The explicitly mapped `Tags` field uses the shared semicolon-separated tag
format. Parsing trims whitespace, ignores empty values, removes duplicates,
and matches or creates tags only in the Clinical catalog. An unmapped or
otherwise unused source column is not interpreted as a tag. Clinical may
display tag names in uppercase; that is Clinical presentation policy, while
shared normalization determines case-insensitive identity.

Import should use explicit preview/validation before commit. Do not silently infer or consolidate separate Placement Opportunities.

Implementation alignment: the current Clinical Sites importer still treats
unmapped columns as tag columns. That legacy behavior conflicts with the
canonical Shared Tag Architecture and is not the approved import contract; the
import UI and parser must require explicit tag-field mappings before this
import path is considered aligned.

## 5. Clinical Tags

Clinical owns its tag catalog. Tags are first-class Clinical module concepts
and are not hardcoded categories. Examples such as `MENTAL HEALTH`,
`ADVANCED`, and `VIRTUAL` are illustrative only. A Clinical tag is identified
by `(clinical-placements, tag_uuid)`, never by name; an identically named tag
in another module is unrelated. Shared identity, normalization, assignment
mechanics, parsing, validation primitives, filtering, and reporting support
are specified in [Shared Tag Architecture](SHARED-TAG-ARCHITECTURE.md).

The **Clinical Sites** tab includes a **Manage Tags** action/dialog. The user may:

- add a tag;
- rename a tag;
- delete a tag.

Rules:

- Clinical displays newly created/renamed tags in uppercase; shared
  normalization provides case-insensitive identity and comparison.
- Renaming a tag must preserve Placement Opportunity associations and all student tag states.
- A tag remains in the Clinical tag registry even if no current Placement Opportunity uses it.
- In Manage Tags, tags not currently used by any Placement Opportunity are visually greyed out. Grey means **unused by a current Clinical Site**, not disabled.
- An unused tag may still be assigned to a student.
- If deletion would affect one or more student Required/Preferred/Restricted selections, warn the user and state how many students are affected before deletion. If confirmed, remove the tag and its Clinical-module associations.

Whenever the tag registry changes, the Student Roster UI must immediately reflect the current tag set.

## 6. Student roster and availability

### 6.1 Student fields and identity

Clinical students are module-owned records/profiles consistent with the product's shared identity boundaries. Imported student data includes:

- student name;
- **student ID (required)**;
- student email;
- recurring day of week;
- availability start time;
- availability end time.

An imported student will normally occupy multiple rows. During a **single import**, rows with the same Student ID represent one Clinical student with multiple recurring Availability Windows.

The user may add, edit, or delete students manually after import, and may edit their availability.

### 6.1.1 Unified Student Roster import with weekly availability (implemented)

Student Roster has a single **Import Student Roster** action (`clinical_students_importer.py`, routes `/api/clinical/students/import/{inspect,preview,apply}`) that creates/updates students and imports their recurring weekly Availability Windows in one operation. There is no standalone availability import. Weekly cell parsing is a pure shared subcomponent (`clinical_availability_importer.py`); persistence is `ClinicalRepository.apply_roster_import`.

The two-phase model applies: Phase A (inspect, map, preview) writes nothing; Phase B requires explicit confirmation, re-reads, re-parses, and re-validates the file against the current roster (never trusting browser preview data), and commits in one transaction.

- Columns: **Student ID** (required), Name, Email, and optional weekday columns Monday–Sunday (Mon–Sun accepted). The user maps them in one dialog.
- Students are matched by Student ID (case-insensitive; the stored ID spelling is kept). Unknown IDs create a student (Name then required); known IDs are updated. Email, when mapped, is authoritative (blank clears it); a blank or unmapped Name keeps the existing name. Duplicate IDs in a file are row errors.
- When at least one weekday column is mapped, each listed student's availability is **replaced**; blank cells mean no availability, so all-blank mapped days clear it. When none is mapped, availability is left unchanged (including for new students, who start with none).
- A weekday cell holds `;`-separated ranges such as `8:00-10:00; T 1:00 PM-4:00 PM`. Entries are Available unless prefixed `T`/`Tentative`; an optional `Available` prefix is accepted. Times use the shared `time_parsing` parser and bare hours (`8`) are accepted. When neither time includes AM/PM and the end time is earlier than the start time, the importer assumes an afternoon end time: `9:00-3:00` means `9:00 AM-3:00 PM`, and `8:00-1:00` means `8:00 AM-1:00 PM`. Existing unmarked ranges beginning before 7:00 (the grid's first hour) continue to be read as afternoon, so `T 1:00-4:00` means 1:00-4:00 PM. Explicit AM/PM values are not adjusted. Same-status overlaps merge; conflicting-status overlaps, malformed times, zero-length ranges, and invalid ranges are row errors.
- Preview shows per row: Student ID, Name, Create/Update, student change (`Create student`, `Update name and email`, `Update name`, `Update email`, `No student change`), and availability change (`Replace 5 availability windows with 8`, `Add N …`, `Clear availability (N windows)`, `No availability change`).
- Any row error rejects the entire import. Roster-only, availability-only, and combined imports each increment `source_data_revision` exactly once and flag an existing schedule stale.

### 6.2 Student roster import is an upsert

Student Roster import is non-destructive. It creates students whose Student ID is new and updates those that match; it never deletes anything.

- Students absent from the file are preserved, along with their availability, Required/Preferred/Restricted tag states, and placements (automatic, manual, and Manually unplaced).
- Clinical Sites, the Clinical Tag registry, and existing placements are preserved; a generated schedule is flagged out of date through the normal source-data revision.
- The confirmation lists the effects above and does not warn of deletion. Removing a student remains an explicit per-student Delete action.

### 6.3 Shared weekly Availability Window editor

The Clinical Student Roster must reuse/generalize the proven Pianist Availability weekly-grid interaction rather than create a separate Clinical-specific interaction model.

The Clinical student availability editor must have the following behavior:

- Show **all seven days, Monday through Sunday**.
- Show times from **7:00 AM through 9:00 PM**.
- Use **30-minute grid cells**, matching the existing Pianist Availability window.
- Support the shared availability states **Available**, **Tentative**, and **Unavailable**.
- Use the same visual language and painting behavior as the Pianist Availability UI wherever practical.
- Support click-and-drag painting. This is mandatory.

For an initially unmarked cell, repeated left-clicks cycle:

`Unmarked -> Available -> Tentative -> Unavailable -> Available -> Tentative -> ...`

A right-click on any cell immediately marks it **Unavailable**. Suppress the normal context menu inside the availability grid when necessary to support this interaction.

Click-and-drag painting must permit efficient marking of ranges. A drag uses a consistent target state for the drag operation rather than independently cycling each cell encountered.

**Closed-world behavior:** unmarked cells are temporary UI state only. When the availability editor is closed/saved, every still-unmarked cell is converted to **Unavailable**. Persisted Clinical availability therefore uses the explicit shared statuses and must not retain a fourth `Unmarked` status.

The 30-minute grid is an **entry/UI resolution**, not a restriction on the shared `AvailabilityWindow` domain primitive. Do not globally change shared availability to assume 30-minute granularity.

On save, contiguous grid cells with the same explicit status should be represented as continuous Availability Windows. Automatic scheduling evaluates these continuous intervals rather than treating cells as discrete appointment slots. Consequently, a clinical session may begin on a 15-minute boundary between the original 30-minute grid boundaries whenever the entire travel-before + session + travel-after commitment is covered by qualifying Available/Tentative intervals. Mixed coverage is legal; for example, a commitment may consume both Available minutes and Tentative minutes.

### 6.4 Tentative availability policy

For automatic Clinical Placement scheduling, availability has the following semantics:

- **Available** time is eligible and preferred.
- **Tentative** time is eligible for automatic scheduling but is less desirable than Available time.
- **Unavailable** time is ineligible for automatic scheduling.

Tentative availability is therefore a soft scheduling preference, not a hard restriction. The optimizer may use Tentative time when necessary to maximize the number of students placed.

The optimizer measures Tentative use as the **total number of Tentative minutes consumed** by the student's complete placement commitment, including travel before, the clinical session itself, and travel after. A placement may span both Available and Tentative intervals; only the Tentative portion contributes Tentative minutes.

Preferred site-tag matching has higher optimization priority than avoidance of Tentative time. Therefore, among schedules placing the same number of students, a solution with more satisfied Preferred student/tag pairs wins even if it consumes more Tentative minutes. Tentative minutes are minimized only after placement count and Preferred-tag satisfaction are tied.

## 7. Student tag requirements, preferences, and restrictions

Selecting a student from the roster exposes the student's availability plus **every tag currently defined in the Clinical tag registry**.

`Required`, `Preferred`, and `Restricted` are Clinical-owned assignment
states, not shared/global tag states. They apply only to Clinical student-tag
assignments; other modules may use different states or no states. The
Student Roster import does not currently import tag assignments. If tag
columns are supported by that import in future, each must be explicitly
mapped to a Clinical state field.

Each student/tag pair has one of four states:

1. Neutral
2. Required
3. Preferred
4. Restricted

Clicking a tag cycles:

`Neutral -> Required -> Preferred -> Restricted -> Neutral`

Visual states:

- Neutral: no color/decoration beyond normal control styling;
- Required: green;
- Preferred: yellow/amber;
- Restricted: red.

Display a compact legend explaining the colors.

Tag semantics for automatic placement:

- Multiple **Required** tags use AND semantics. A candidate Placement Opportunity must contain every Required tag.
- If a candidate contains **any Restricted** tag, it is ineligible for automatic placement.
- Preferred tags are cumulative and equally weighted. Each satisfied **student/tag pair** contributes exactly one Preferred-match unit to the optimization objective. A student's Preferred tag contributes at most once to that student's placement, regardless of duplicated source data. A candidate satisfying two Preferred tags therefore contributes two units, while one satisfying a single Preferred tag contributes one.

## 8. Automatic placement optimizer contract

The Clinical optimizer is module-specific. Do not reuse or coerce the Accompanist or Jury optimizer.

### 8.1 Student placement cardinality

Each student receives at most **one** Clinical placement. A student may remain unplaced if no placement can be made under the approved automatic rules.

### 8.2 Candidate eligibility

An automatic candidate placement is valid only when all of the following are true:

- the Placement Opportunity contains every Required tag for the student;
- the Placement Opportunity contains none of the student's Restricted tags;
- the student has sufficient qualifying Available and/or Tentative availability on the same recurring weekday;
- the student's availability continuously accommodates **travel before + entire clinical session + travel after**. The commitment may cross adjacent Available/Tentative windows, and any Tentative portions contribute Tentative minutes to the optimization objective;
- the Placement Opportunity has unused automatic capacity after accounting for preserved manual placements. For generation, remaining capacity is `max(0, capacity - preserved manual assignments)`. If manual assignments already equal or exceed capacity, preserve them, show the applicable capacity warning, and make no additional automatic assignments to that opportunity.

Travel time occurs before and after the session but does not change the site's clinical session start/end. Example: a 60-minute session with 15-minute travel each way consumes 90 continuous minutes of the student's relevant availability.

### 8.3 Automatic start times

The optimizer chooses the student's exact clinical-session start time.

- Candidate automatic start times occur at **15-minute increments**.
- The clinical session itself must fit within the Placement Opportunity's start/end window.
- The student's surrounding availability must also cover travel before and travel after.
- If otherwise equivalent assignments/times remain, choose the **earliest possible session start**.

The 15-minute optimizer increment is intentionally independent of the 30-minute manual availability-grid resolution.

### 8.4 Optimization priorities

Use lexicographic objectives, not an arbitrary blended score:

1. **Maximize the number of students placed.**
2. Subject to objective 1, **maximize the total number of satisfied Preferred student/tag pairs**.
3. Subject to objectives 1 and 2, **minimize the total Tentative availability minutes consumed**, including travel before, the clinical session, and travel after.
4. Subject to objectives 1 through 3, **prefer earlier clinical-session start times** where choices are otherwise equivalent.
5. Resolve any remaining ties deterministically using stable, clinically neutral ordering. Do not invent clinical significance for the final tie-break.

A schedule placing more students always wins. Among schedules placing the same number of students, Preferred-tag satisfaction outranks Tentative-time avoidance. Tentative-time avoidance outranks earlier start times.

Required, Restricted, Unavailable time, Placement Opportunity time-window, and capacity rules are hard constraints for **automatic** placement rather than weighted penalties. Do not automatically violate one in order to place another student.

The solver must optimize globally enough to protect hard-to-place students. Do not use a naive roster-order greedy algorithm if it can unnecessarily leave a student unplaced because a flexible student consumed that student's only legal opportunity.

## 9. Schedule lifecycle, stale state, and clearing

The Schedule tab displays every student and their current placement state.

Primary actions include:

- **Generate Placements**
- **Delete All Placements**

### 9.1 Placement states

At minimum distinguish:

- ordinary **Unplaced**: default state; visually amber; eligible for automatic generation;
- automatically generated placement;
- manually assigned/edited placement;
- **Manually unplaced**: explicit user choice; visually amber; excluded from subsequent automatic generation.

Manual placements and Manually unplaced decisions must survive subsequent Generate Placements operations. The optimizer schedules remaining eligible students around those manual decisions and treats capacity consumed by preserved manual placements as occupied for automatic generation.

Clicking **Generate Placements** again discards/replaces prior **automatically generated** Clinical placements before generating a fresh automatic solution. It does not discard manual placements or Manually unplaced decisions.

### 9.2 Delete All Placements

`Delete All Placements` requires confirmation and then clears both generated and manual Clinical placement decisions, including Manually unplaced state. It does not delete student or Clinical Site source data. All students return to ordinary Unplaced. Clearing all placements also clears generated-result revision/staleness state because no generated schedule remains to be stale.

### 9.3 Source-data changes and stale schedule

Once placements exist, edits to scheduling source data must **not silently delete or rewrite the schedule**.

If relevant source data changes, preserve the visible schedule and mark it prominently as out of date, for example:

> **Schedule is out of date.** Student or Clinical Site data has changed since placements were last generated. Existing placements have been preserved. Review warnings, regenerate placements, or delete the schedule.

Relevant changes include student availability, student tag states, Placement Opportunity windows, capacity, travel time, session length, tags, and adding/deleting opportunities or students as applicable.

Revalidate retained placements against current source data and show current warnings. Staleness and a placement-specific warning are separate concepts and may appear simultaneously.

**Deleting an in-use Placement Opportunity is a special destructive case.** If one or more students are assigned to the opportunity, require confirmation and state how many placements will be affected. On confirmation, delete the opportunity, remove only those affected assignments, return those students to ordinary Unplaced, preserve all other placements, and mark the remaining schedule out of date. Do not retain dangling placement references to a deleted opportunity.

## 10. Schedule UI and manual override

Use a compact student-oriented schedule table conceptually containing:

- Student
- Placement Opportunity
- Day
- Session time
- Status/warnings

The Placement Opportunity selector must identify the actual opportunity rather than show only an ambiguous site name. A user-facing option should include enough context, for example:

`Mercy Hospital | Tue 8:00 AM-11:00 AM`

Do not expose internal IDs.

For an automatically **Unplaced** student, provide useful diagnostic reasons where they can be determined, such as no opportunity satisfying all Required tags, Restricted-tag conflicts, lack of qualifying availability, or otherwise eligible opportunities being at capacity. Diagnostics should help the coordinator understand why generation failed without pretending to be a formal proof of global infeasibility.

### 10.1 Manual start time

Automatic start times are 15-minute aligned, but **manual times are not**.

When manually assigning/editing a placement:

- allow the user to type a start time directly;
- use **12-hour time with AM/PM** in the UI;
- do not force 15-minute increments;
- a time such as `8:10 AM` is valid manual input;
- accept ordinary case-insensitive 12-hour forms such as `8:10 AM`, `8:10 am`, and `8:10am`, then normalize the displayed value;
- reject ambiguous, incomplete, or invalid clock inputs with inline validation rather than silently guessing;
- do not require the user to enter an end time;
- calculate end time as `manual start + Placement Opportunity session length`.

Manual start times may extend outside student availability or the Placement Opportunity window. Permit the edit and show warnings rather than blocking it.

### Three-State Placement Model (Unassigned / Assigned / Manually Unplaced)

Clinical Placement and Lesson Scheduling share the same placement-state model. The API exposes it as the typed `placement_state` field (`unassigned`, `assigned`, `manuallyUnplaced`) on every schedule entry. It is derived, not stored: Unassigned is the absence of a placement record.

| State | Placement | Generate Schedule |
|---|---|---|
| Unassigned | None | Eligible; may be assigned automatically |
| Assigned (Automatic or Manual) | Exists | Automatic is replaceable; Manual is preserved |
| Manually Unplaced | None (explicit coordinator decision) | Excluded; never assigned |

Placement dropdown order is always: `Unassigned`, a separator, the Placement Opportunity sites, a separator, `Manually Unplaced`.

Example:

```
Unassigned
────────────
Mercy Behavioral Health
Campus Counseling Center
Hope Community Clinic
────────────
Manually Unplaced
```

- Choosing Unassigned removes any placement or Manually Unplaced decision (`DELETE /api/clinical/schedule/placements/{student_id}`) and returns the student to the automatic pool. It does not regenerate.
- Choosing a placement option saves a manual placement; choosing Manually Unplaced saves the exclusion decision.
- The summary reports Unassigned and Manually Unplaced counts separately.
- Both modules use identical terminology; this is the same model in both specifications.

### 10.2 Manual override principle

The automatic optimizer is strict; the human coordinator is authoritative.

The user may manually place any student at any Placement Opportunity even when the assignment violates automatic constraints. Do not block the action. Validate it and display all applicable warnings independently, including at least:

- outside student availability;
- missing Required tag (name the tag where practical);
- Restricted tag conflict (name the tag where practical);
- Placement Opportunity capacity exceeded;
- manual session outside the Placement Opportunity time window.

If several problems apply, show all relevant warnings rather than collapsing them into a generic `invalid placement` message.

Manually exceeding capacity is allowed. Example warning:

`Capacity exceeded: 3 students assigned; capacity is 2.`

## 11. Reports

The rightmost **Reports** tab follows the established reporting pattern of the other modules and uses shared report rendering/export infrastructure without adopting another module's report schema.

Provide a printable Markdown-format report containing at least two sections:

1. **Student-oriented Placement Roster**
2. **Site-oriented Placement Roster**

The report should include enough information to be operationally useful, including relevant student identity/contact, Placement Opportunity/site/contact, day, session start/end, and status where appropriate. Include unplaced students and meaningful warnings in an appropriate section rather than silently omitting unfinished work.

The student-oriented roster should use a deterministic student-name ordering.

The site-oriented roster should use deterministic ordering/grouping by site name, then Placement Opportunity weekday, then Placement Opportunity start time, then assigned student session start, then student name. This allows the coordinator to see which students are assigned to each site/opportunity and their session times.

Provide buttons/actions to export locally as:

- `.md`
- `.pdf`

Follow the existing Accompanist/Jury reporting mechanics and local-only privacy behavior.

## 12. Persistence, revisions, and result ownership

Clinical module data belongs to the active Scheduling Session but remains module-owned.

Persist enough revision/provenance information to determine whether an existing Clinical schedule is stale relative to Clinical scheduling inputs. Do not tie Clinical staleness to Accompanist or Jury revisions because there are no cross-module dependencies for this workflow.

Use versioned schema migrations for new persistence. Do not create ad hoc schema mutations. Keep `.mpsession` compatibility/versioning rules and local recovery behavior intact.

A Clinical placement result should use a typed Clinical-owned result contract/payload rather than a universal schedule/assignment object.

## 13. Validation and time semantics

- Site and student weekday windows recur weekly for the term.
- No individual-date exception model is required for this module at this stage.
- UI clock times should be presented in **12-hour AM/PM** format.
- Use established local-time primitives; no session timezone requirement is introduced.
- Treat intervals consistently with the application's shared half-open time-window conventions where appropriate.
- A Placement Opportunity must have a valid positive session length and capacity.
- Capacity omitted on import defaults to 1.
- Invalid/malformed imports should surface structured validation issues and must not silently invent data.

## 14. Required test coverage

Add module-specific automated tests at the pure-domain/service level before depending on UI tests. At minimum cover:

### Import and tags

- multiple rows sharing a site name remain distinct Placement Opportunities;
- capacity default = 1;
- only explicitly mapped Clinical tag fields are parsed; unmapped columns never create tags;
- semicolon-separated mapped values, blank values, duplicate removal, and preview behavior follow the Shared Tag Architecture;
- Clinical uppercase display and shared case-insensitive tag identity;
- duplicate normalized tags collapse;
- deleting a student-used tag requires warning/confirmation behavior;
- roster import upserts by Student ID (preserving absent students, tag states, and placements) and optionally replaces listed students' availability.

### Availability editor/model

- Monday-Sunday and 7:00 AM-9:00 PM grid bounds;
- 30-minute grid resolution;
- click state cycle;
- right-click to Unavailable;
- click-and-drag paint behavior;
- unmarked cells become Unavailable on close/save;
- shared availability domain remains capable of non-30-minute intervals.

### Automatic placement

- Required AND semantics;
- any Restricted tag excludes an automatic candidate;
- cumulative Preferred student/tag pairs form the secondary optimization objective;
- placement count outranks Preferred-match count;
- travel is required before and after the session;
- exact boundary cases for travel/session/site/student windows;
- 15-minute automatic start generation;
- earliest-time tie-break;
- total opportunity capacity, independent of overlapping student session times;
- global assignment cases where greedy roster ordering would lose a placement;
- deterministic output for identical input;
- Tentative remains eligible when its use increases the number of students placed;
- Preferred-tag matches outrank minimizing Tentative minutes when placement count is equal;
- Tentative usage is measured in minutes across travel-before, session, and travel-after;
- mixed Available/Tentative coverage is eligible and counts only its Tentative portion;
- 15-minute session starts can be valid inside continuous windows derived from the 30-minute entry grid;
- Unavailable remains ineligible for automatic placement.

### Manual placement and lifecycle

- arbitrary minute manual start such as 8:10 AM;
- calculated session end;
- warnings for student availability, Required, Restricted, capacity, and site-window violations;
- manual violations are allowed rather than blocked;
- manual placements survive regeneration;
- prior automatic placements are discarded/regenerated on Generate Placements;
- manual over-capacity assignments yield zero remaining automatic capacity without deleting the manual assignments;
- Manually unplaced students remain excluded from regeneration;
- ordinary Unplaced students remain eligible;
- source-data changes preserve placements and set stale state;
- current warnings recalculate after source edits;
- Delete All Placements clears generated/manual/manually-unplaced decisions but not source data and clears stale/result revision state;
- Student Roster replacement clears all Clinical placements while preserving Clinical Site data;
- deleting an in-use Placement Opportunity requires confirmation, unplaces only affected students, and leaves no dangling placement reference;
- useful diagnostics are available for automatically unplaced students.

### Reports and privacy

- student-oriented and site-oriented report sections and deterministic ordering;
- unplaced/warning visibility;
- Markdown/PDF output remains local;
- no network dependency is introduced.

## 15. Suggested implementation sequence

Keep commits/phases narrow and testable. A sensible sequence is:

1. **Clinical persistence and migrations**: module-owned entities, tag registry, Placement Opportunities, students, student-tag state, Clinical placement result/revision state.
2. **Clinical Sites UI/import**: flat Placement Opportunity list, add/edit/delete, CSV/XLSX mapping/preview, capacity, explicitly mapped Clinical tag import, Manage Tags.
3. **Student Roster UI/import**: destructive replacement import with warning; manual student editing.
4. **Shared availability reuse/generalization**: adapt the existing Pianist Availability grid for Clinical students while preserving Pianist behavior; Monday-Sunday, 7 AM-9 PM, 30-minute cells, three-state cycle, right-click, drag painting, closed-world save.
5. **Clinical validation services**: candidate eligibility, travel/session feasibility, capacity, tags, manual warning calculation, stale-input revision tracking.
6. **Clinical optimizer**: pure deterministic domain implementation; lexicographic placement/preference/Tentative objectives; 15-minute candidate starts; tests before UI integration.
7. **Schedule UI**: Generate Placements, preserved manual assignments, Unplaced/Manually unplaced, free-form 12-hour manual start editing, warnings, Delete All Placements, stale schedule banner.
8. **Reports**: Markdown preview/representation, student-oriented and site-oriented sections, `.md` and `.pdf` export.
9. **Regression and integration pass**: confirm Accompanist and Jury remain behaviorally unchanged and all existing tests continue to pass.

Do not combine these phases into a large rewrite. Prefer extending existing module registration, shared import/availability/report infrastructure, migration conventions, and result-envelope patterns.

## 16. Explicit non-goals

Do not add any of the following as part of this module unless separately approved:

- cross-module data exchange between Clinical, Accompanist, or Jury;
- cloud storage or server-side scheduling;
- Forms/Graph API integration;
- telemetry or analytics;
- a universal optimizer or universal Schedule/Assignment model;
- automatic consolidation of same-name Clinical Sites;
- simultaneous-occupancy scheduling constraints beyond the stated total capacity;
- date-specific exceptions/holidays;
- weighted Preferred tags or strong/weak preference levels;
- automatic violation of Required/Restricted/time/capacity constraints to increase placement count;
- preservation/reconciliation of student tag settings across Student Roster replacement imports;
- exposing internal record IDs to users.

## 17. Canonical architectural alignment

This specification is subordinate to the repository's canonical product/platform architecture documents for concerns outside the Clinical module. In particular:

- `MODULE-ARCHITECTURE.md` defines module boundaries and rejects universal optimizer/schema designs.
- `PRODUCT-MODULE-ARCHITECTURE.md` defines Scheduling Session ownership, shared availability/import/reporting infrastructure, versioned migrations/results, and local execution architecture.
- `PRIVACY-ARCHITECTURE.md` requires scheduling data and imported workbooks to remain local and prohibits telemetry, analytics, cloud storage, and external network integration.

Where this document is more specific about Clinical Placement workflow and domain semantics, it is the implementation specification for the Clinical module. Where a future change conflicts with product/session/privacy architecture, update the architectural decision explicitly rather than coding around it.
