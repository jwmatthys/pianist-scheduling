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

Any additional imported columns are **tag-value columns**. Their headings have no semantic meaning after mapping the standard fields.

Tag import rules:

- Each nonblank cell in an additional column contributes exactly **one tag** to that Placement Opportunity.
- Blank cells contribute no tag.
- Do not split a cell on commas, semicolons, or other delimiters. One cell equals one tag.
- Trim surrounding whitespace.
- Tag identity is case-insensitive.
- Normalize tag display text to **UPPERCASE** when tags are created/imported.
- Duplicate normalized tags on the same Placement Opportunity collapse to one association.

Import should use explicit preview/validation before commit. Do not silently infer or consolidate separate Placement Opportunities.

## 5. Clinical Tags

Tags are first-class Clinical module concepts and are not hardcoded categories. Examples such as `MENTAL HEALTH`, `ADVANCED`, and `VIRTUAL` are illustrative only.

The **Clinical Sites** tab includes a **Manage Tags** action/dialog. The user may:

- add a tag;
- rename a tag;
- delete a tag.

Rules:

- Normalize newly created/renamed tags to uppercase.
- Compare tag identity case-insensitively.
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

### 6.2 Student roster re-import is full replacement

Clinical Student Roster import is deliberately destructive and follows the replacement behavior used elsewhere in the product:

- Before commit, explicitly warn that importing will delete/replace the current Clinical student roster.
- On confirmation, clear the existing Clinical student data for this module and replace it with the imported roster.
- Do **not** preserve student tag states by Student ID.
- Existing Required, Preferred, and Restricted selections are cleared as part of roster replacement.
- Clear **all existing Clinical placements**, including automatic placements, manual placements, and Manually unplaced decisions. Placements belonging to the roster being replaced must not survive as dangling schedule records.
- Do not attempt to reconcile the new roster against the previous one.

The confirmation must explicitly state that replacing the roster deletes the current Clinical student roster, student availability, student Required/Preferred/Restricted tag selections, and all Clinical placements. Clinical Site/Placement Opportunity data is not changed.

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
- nonstandard column values create tags while headers do not;
- blank extra cells create no tag;
- one cell remains one tag (no delimiter splitting);
- whitespace trimming and uppercase/case-insensitive tag normalization;
- duplicate normalized tags collapse;
- deleting a student-used tag requires warning/confirmation behavior;
- roster re-import fully replaces students and clears tag states.

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
2. **Clinical Sites UI/import**: flat Placement Opportunity list, add/edit/delete, CSV/XLSX mapping/preview, capacity, tag-value import, Manage Tags.
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
