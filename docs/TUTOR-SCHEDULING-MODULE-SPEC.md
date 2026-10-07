# Tutor Scheduling Module Overview and Implementation Specification

**Status:** Draft module design for implementation.  
**Product:** Music Program Scheduler  
**Module:** Tutor Scheduling

## 1. Purpose and implementation guardrails

The Tutor Scheduling module assigns a roster of tutoring Clients to Tutors based on recurring weekly availability, client-specific session lengths, Tutor capacity, and client-defined Required, Preferred, and Restricted tags.

Tutor Scheduling is a completely independent product module. It may reuse proven infrastructure, UI components, and a shared constrained-assignment optimization mechanism, but it must **not share domain data with any other scheduling module**.

Even when identifying information happens to be identical, records in different modules are independent. For example, a person with the same name, email, or ID in Tutor Scheduling and Clinical Placement Scheduling is not a shared application entity. Editing, deleting, importing, or scheduling that person in one module must have no effect on the other module.

During implementation:

- Do not create cross-module foreign keys, identity reconciliation, synchronization, or implicit data reuse.
- Do not import Clients from Accompanist, Jury, or Clinical data.
- Do not import Tutors from Pianist or other module data.
- Keep Tutor and Client records, tag selections, availability, assignments, revisions, imports, and reports module-owned.
- Shared code must not imply shared persisted records.
- Reuse/generalize proven shared infrastructure where appropriate instead of cloning behavior.
- A shared constrained-assignment optimizer may be used by Tutor Scheduling and Clinical Placement Scheduling, but module-specific candidate generation and validation remain module-owned.
- Do not force Accompanist Scheduling or Jury Scheduling through this optimizer abstraction.
- Scheduling/imported data remains local to the user's device. Do not add cloud services, telemetry, analytics, or network integrations.

## 2. Module navigation and workflow

The Tutor Scheduling module has four tabs:

1. **Tutors**
2. **Client Roster**
3. **Schedule**
4. **Reports**

Expected workflow:

1. Enter or import Tutors.
2. Review/edit Tutor information, availability, capacity, and tags.
3. Manage the Tutor Scheduling tag registry.
4. Enter or import the Client roster and client availability.
5. Review/edit Clients, client-specific session lengths, availability, and tag requirements/preferences/restrictions.
6. Generate the tutoring schedule.
7. Review unplaced Clients and warnings; make manual Tutor assignments or time adjustments as needed.
8. Generate/review reports and export Markdown or PDF.

## 3. Shared concepts versus module-owned data

Tutor Scheduling is intentionally similar to Clinical Placement Scheduling, but similarity of workflow does not create shared data.

### 3.1 Appropriate shared mechanisms

The implementation should reuse or generalize, where appropriate:

- the weekly Availability Window editor;
- shared availability primitives and status vocabulary;
- CSV/XLSX import/mapping/preview infrastructure;
- tag normalization and reusable tag-control UI patterns;
- Required/Preferred/Restricted interaction patterns;
- report rendering/export infrastructure;
- revision/staleness patterns;
- a small pure constrained-assignment optimizer operating on normalized candidate assignments.

### 3.2 Module-owned concepts

Tutor Scheduling owns its own:

- Tutors;
- Clients;
- Tutor availability;
- Client availability;
- Tutor capacities;
- tag registry and tag associations;
- Client Required/Preferred/Restricted states;
- tutoring session lengths;
- generated/manual assignments;
- validation/warning semantics;
- import schemas;
- result/revision records;
- reports and display terminology.

Do not create a universal persisted Student, Person, Tutor/Pianist, Schedule, or Assignment record merely because fields overlap with another module.

## 4. Tutors

### 4.1 Tutor fields

Each Tutor should contain at least:

- internal application-generated identifier, never required from or displayed to the user;
- Tutor name;
- Tutor ID where provided/used by the import schema;
- Tutor email;
- recurring weekly availability;
- capacity;
- zero or more tags.

Capacity defaults to **1** when omitted on import or manual creation unless the user supplies another value.

Tutor capacity means the **maximum total number of Clients assigned to that Tutor**, regardless of whether tutoring sessions overlap in time. It is not a simultaneous-occupancy limit.

Example: Tutor capacity 4 permits no more than four automatically assigned Clients in total. Those Clients' scheduled tutoring times may overlap completely. Tutor Scheduling does not introduce a conflict rule merely because two Clients are scheduled with the same Tutor at the same time.

### 4.2 Tutor editing

After import, Tutor records are ordinary editable application data. The user may:

- add a Tutor;
- edit Tutor identity/contact fields;
- edit availability;
- edit capacity;
- add/remove Tutor tags;
- delete a Tutor.

Edits should not silently destroy unaffected assignments. Schedule staleness and destructive deletion behavior are defined below.

## 5. Weekly Availability Window editor

Both Tutors and Clients use the reusable weekly Availability Window editor based on the proven Pianist Availability interaction.

Requirements:

- Show **Monday through Sunday**.
- Show **7:00 AM through 9:00 PM**.
- Use **30-minute grid cells**.
- Support **Available**, **Tentative**, and **Unavailable**.
- Support click-and-drag painting. This is mandatory.
- Reuse/generalize the existing component and painting logic rather than implementing separate Tutor and Client grids.

For an initially unmarked cell, repeated left-clicks cycle:

`Unmarked -> Available -> Tentative -> Unavailable -> Available -> ...`

A right-click immediately marks a cell **Unavailable**. Suppress the normal context menu inside the grid where necessary.

A drag operation paints a consistent target state across the traversed cells rather than independently cycling each cell.

**Closed-world save behavior:** any cells still unmarked when the editor closes/saves become **Unavailable**. `Unmarked` is a temporary editing condition, not a persisted availability status.

The 30-minute grid is an entry/UI resolution only. Persist availability through the application's shared continuous Availability Window model rather than treating the grid cells as appointment slots. Contiguous cells with the same status should be represented as continuous recurring weekly windows.

Because automatic tutoring sessions may begin at 15-minute boundaries, a session start may fall between the original 30-minute grid boundaries whenever its entire duration is continuously covered by qualifying availability for both Client and Tutor.

## 6. Tutor Scheduling tags

Tags are first-class Tutor Scheduling concepts and are not hardcoded categories. They are completely independent of Clinical Placement tags, even when their display text is identical.

The **Tutors** tab includes **Manage Tags**. The user may:

- add a tag;
- rename a tag;
- delete a tag.

Rules:

- Trim surrounding whitespace.
- Tag identity is case-insensitive within Tutor Scheduling.
- Normalize tag display values to **UPPERCASE**.
- Renaming a tag preserves Tutor associations and Client tag states within this module.
- Tags remain in the registry when no Tutor currently uses them.
- Tags unused by all current Tutors appear greyed out in Manage Tags. Grey means unused, not disabled.
- An unused tag may still be assigned to a Client.
- Deleting a tag used as Required, Preferred, or Restricted by one or more Clients requires confirmation stating how many Clients are affected.

No Tutor Scheduling tag creation, rename, or deletion affects another module's tags.

## 7. Tutor import

Accept local CSV and XLSX through the application's existing local import/mapping/preview infrastructure.

Tutor import should support Tutor identity/contact fields, capacity, recurring availability, and tag values. The exact user-facing import template should follow established mapping/preview conventions rather than relying on column order.

Where one Tutor occupies multiple source rows because of multiple availability windows, consolidate those rows only according to the Tutor Scheduling import identity policy. Do not reconcile imported Tutors against records from another module.

Additional tag-value columns follow the same Tutor Scheduling tag rules:

- column headings have no semantic meaning as tags;
- each nonblank additional cell contributes one tag;
- one cell equals one tag;
- blank cells contribute no tag;
- do not split cell contents on commas/semicolons;
- normalize and deduplicate tags case-insensitively within Tutor Scheduling.

Import/re-import should use explicit preview, validation, warning, and replacement semantics consistent with the application's established module-owned import behavior. Do not perform clever cross-module merges.

## 8. Client roster

### 8.1 Client fields

Each Tutor Scheduling Client contains:

- internal application-generated identifier;
- Client name;
- **Client ID (required)**;
- Client email;
- recurring weekly availability;
- **session length**, client-specific and defaulting to **60 minutes**;
- a state for every Tutor Scheduling tag;
- zero or one Tutor assignment.

Client records are Tutor Scheduling records only. A matching Client/Student ID in Clinical, Jury, or Accompanist data does not create a relationship.

The user may add, edit, and delete Clients manually after import, including availability and session length.

### 8.2 Client Roster import

Imported Client data should support:

- Client name;
- Client ID;
- Client email;
- recurring day of week;
- availability start time;
- availability end time;
- session length where supplied.

Rows sharing the same Client ID within a single Tutor Scheduling import represent one Client with multiple recurring Availability Windows. If session length is omitted, default it to 60 minutes.

Re-import is a **full replacement** of the Tutor Scheduling Client roster:

- warn before commit;
- delete/replace existing Tutor Scheduling Clients and Client availability;
- clear existing Tutor Scheduling Client tag states;
- clear all existing Tutor Scheduling assignments, including generated assignments, manual assignments, and Manually unplaced decisions;
- do not preserve or reconcile Client data by ID;
- do not change Tutor records, Tutor availability, Tutor tags, or capacities;
- do not affect any other module.

The confirmation must clearly identify all Tutor Scheduling data that will be lost.

## 9. Client tag states

Selecting a Client exposes the Client's availability, session length, and every currently defined Tutor Scheduling tag.

Each Client/tag pair has four states:

1. Neutral
2. Required
3. Preferred
4. Restricted

Click cycle:

`Neutral -> Required -> Preferred -> Restricted -> Neutral`

Visual states:

- Neutral: default/plain;
- Required: green;
- Preferred: yellow/amber;
- Restricted: red.

Display a compact legend.

Automatic matching semantics:

- Multiple Required tags use **AND** semantics. A Tutor must possess every Required tag.
- Any Tutor tag matching a Client Restricted tag makes that Tutor ineligible for automatic assignment to that Client.
- Preferred tags are cumulative and equally weighted.
- Each satisfied **Client/tag pair** contributes exactly one Preferred-match unit, at most once for that Client assignment.

## 10. Automatic Tutor scheduling contract

Each Client receives at most **one Tutor assignment**, with the primary goal of assigning as many Clients as possible.

### 10.1 Automatic candidate eligibility

A Tutor/Client/start-time candidate is eligible only when:

- the Tutor has every Client Required tag;
- the Tutor has none of the Client Restricted tags;
- Client and Tutor have continuously qualifying Available and/or Tentative availability for the Client's complete session length on the same recurring weekday;
- the exact tutoring session fits entirely inside both participants' qualifying availability;
- the Tutor has unused automatic capacity after preserved manual assignments are accounted for.

`Unavailable` time on either side makes the candidate ineligible.

There is no travel-time requirement in Tutor Scheduling.

### 10.2 Automatic start times

The optimizer chooses the tutoring session's exact weekly start time.

- Candidate automatic starts occur at **15-minute increments**.
- The Client's full session duration must fit within qualifying Client and Tutor availability.
- A 15-minute start may occur between the 30-minute boundaries used by the availability-entry grid.
- Earlier starts are preferred only after all higher optimization objectives are tied.

### 10.3 Tentative-minute calculation

Tentative time is a soft cost on **both sides** of a Tutor assignment and has equal value for Client and Tutor.

For each candidate, calculate:

`total tentative minutes = Client tentative minutes during session + Tutor tentative minutes during session`

Examples:

- 60-minute session, Client entirely Available, Tutor entirely Available -> 0 Tentative minutes.
- 60-minute session, Client has 30 Tentative minutes, Tutor entirely Available -> 30 Tentative minutes.
- 60-minute session, Client has 30 Tentative minutes and Tutor has 15 Tentative minutes -> 45 Tentative minutes.
- 60-minute session entirely Tentative for both -> 120 Tentative minutes.

A session may cross adjacent Available/Tentative windows; only portions marked Tentative contribute to this objective.

### 10.4 Capacity during generation

Remaining automatic capacity is:

`max(0, Tutor capacity - preserved manual assignments)`

If preserved manual assignments already equal or exceed capacity:

- preserve those manual assignments;
- display the relevant capacity warning;
- make no additional automatic assignments to that Tutor.

Tutor capacity remains a total-assignment cap, not a time-overlap constraint.

### 10.5 Optimization hierarchy

Use lexicographic objectives, not an arbitrary blended score:

1. **Maximize the number of Clients assigned to Tutors.**
2. Subject to #1, **maximize the total number of satisfied Preferred Client/tag pairs.**
3. Subject to #1-2, **minimize total Tentative availability minutes consumed across Clients and Tutors.**
4. Subject to #1-3, **prefer earlier tutoring-session start times.**
5. Resolve remaining ties using stable, deterministic, semantically neutral ordering.

A solution assigning more Clients always wins. Among solutions assigning the same number of Clients, Preferred-tag satisfaction outranks Tentative-time avoidance. Tentative-time avoidance outranks earlier start times.

Required tags, Restricted tags, Unavailable time, and automatic capacity are hard automatic constraints rather than weighted penalties.

The algorithm must optimize globally enough to protect difficult-to-place Clients. Do not use a naive roster-order greedy algorithm when doing so can unnecessarily leave a Client unassigned.

## 11. Shared constrained-assignment optimizer boundary

Clinical Placement Scheduling and Tutor Scheduling may share a **small pure constrained-assignment optimization engine**, because they use the same lexicographic assignment/capacity objective structure. This does not make their domain models shared.

The preferred boundary is:

1. Module-specific code produces normalized legal candidate assignments.
2. Each candidate exposes only the generic information needed for optimization, conceptually including:
   - consumer key;
   - provider key;
   - candidate start time;
   - Preferred-match count;
   - Tentative-minute cost.
3. Shared optimization selects candidates subject to:
   - at most one candidate per consumer;
   - provider total capacity;
   - preserved/manual capacity consumption;
   - the lexicographic objectives.
4. Module-specific code converts the selected result into a typed module-owned result and computes module-specific warnings/display data.

The shared optimizer should not need to know what a `Tutor`, `Client`, `Clinical Student`, or `Placement Opportunity` is. Do not put module switches such as `if clinical ... else if tutor ...` into the shared optimization core.

Tutor-specific candidate generation remains responsible for Tutor/Client overlapping availability and Client session length. Clinical-specific candidate generation remains responsible for its own site-window, travel-time, and Clinical rules.

## 12. Schedule lifecycle

The Schedule tab displays every Client and current assignment state.

Primary actions:

- **Generate Schedule** (or **Generate Tutor Assignments**, if that matches final product terminology)
- **Delete All Assignments**

At minimum distinguish:

- ordinary **Unplaced**: default state, amber, eligible for automatic generation;
- automatically generated Tutor assignment;
- manually assigned/edited Tutor assignment;
- **Manually unplaced**: deliberate human decision, amber, excluded from subsequent automatic generation.

### 12.1 Regeneration

Manual assignments and Manually unplaced decisions survive subsequent generation.

When generation runs again:

- preserve manual assignments;
- preserve Manually unplaced decisions;
- discard/replace prior automatically generated assignments;
- calculate remaining Tutor capacities after preserved manual assignments;
- optimize all remaining eligible Clients.

### 12.2 Delete All Assignments

Require confirmation. On confirmation:

- clear automatic assignments;
- clear manual assignments;
- clear Manually unplaced decisions;
- return all Clients to ordinary Unplaced;
- leave Tutors, Clients, availability, capacities, tags, and other source data intact;
- clear generated-result revision/staleness state because no generated schedule remains.

## 13. Schedule staleness and source-data edits

Once assignments exist, ordinary edits to relevant Tutor Scheduling source data do not silently destroy unaffected assignments.

Relevant changes include:

- Client availability;
- Tutor availability;
- Client session length;
- Client tag states;
- Tutor tags;
- Tutor capacity;
- adding/editing relevant Tutor or Client scheduling data.

Preserve the visible schedule and show a prominent stale warning, such as:

> **Schedule is out of date.** Tutor or Client data has changed since the schedule was last generated. Existing assignments have been preserved. Review warnings, regenerate the schedule, or delete all assignments.

Revalidate preserved assignments against current source data and present current warnings.

### 13.1 Deleting an assigned Tutor

Deleting a Tutor with one or more assigned Clients is a special destructive case:

- require confirmation;
- state how many assignments are affected;
- on confirmation, delete the Tutor;
- remove only assignments to that Tutor;
- return affected Clients to ordinary Unplaced;
- preserve unrelated assignments;
- mark the remaining schedule out of date;
- leave no dangling assignment references.

### 13.2 Deleting an assigned Client

Deleting a Client removes that Client's assignment after confirmation where appropriate. It does not modify the assigned Tutor or unrelated Clients and Tutors. If other assignments remain, mark the schedule out of date.

## 14. Schedule UI and manual override

Use a compact Client-oriented schedule conceptually showing:

- Client
- Tutor
- Day
- Session time
- Status/warnings

The Tutor selector contains all Tutors plus:

- Unplaced
- Manually unplaced

Do not expose internal record IDs.

### 14.1 Manual start time

Automatic start times use 15-minute increments; manual times do not.

For manual assignment/editing:

- allow a start time to be typed directly;
- display 12-hour time with AM/PM;
- accept normal case-insensitive forms such as `8:10 AM`, `8:10 am`, and `8:10am`;
- normalize the displayed value;
- reject ambiguous/incomplete/invalid times with inline validation rather than guessing;
- do not require an end time;
- calculate end time as `manual start + Client session length`.

A manual start may violate Client or Tutor availability. Permit the assignment and warn rather than blocking it.

### 14.2 Human override principle

Automatic scheduling is strict; the human scheduler is authoritative.

The user may manually assign any Client to any Tutor and may manually enter a time even when automatic constraints fail. Show all applicable warnings independently rather than collapsing them into a generic invalid state.

Warnings include at least:

- outside Client availability;
- outside Tutor availability;
- missing Required tag;
- Restricted tag conflict;
- Tutor capacity exceeded.

Manual over-capacity assignment is allowed. Example:

`Capacity exceeded: 5 Clients assigned; Tutor capacity is 4.`

If a manual session uses Tentative availability, that may be visibly indicated where useful, but Tentative is not itself an invalid manual assignment.

### 14.3 Automatically unplaced diagnostics

For automatically Unplaced Clients, provide useful diagnostic reasons where determinable, such as:

- no Tutor satisfies all Required tags;
- otherwise possible Tutors conflict with Restricted tags;
- no Client/Tutor qualifying availability overlap supports the required session length;
- otherwise eligible Tutors are at capacity.

Diagnostics should assist the user without claiming to be a formal proof of global infeasibility.

## 15. Reports

The rightmost **Reports** tab follows the reporting pattern established elsewhere in the application while retaining Tutor Scheduling terminology and module-owned report fields.

Provide a printable Markdown-format report with at least:

1. **Client-oriented Tutoring Roster**
2. **Tutor-oriented Tutoring Roster**

### 15.1 Client-oriented roster

Show each Client's assignment and operationally useful details, including as appropriate:

- Client name/ID/email;
- Tutor name/contact;
- recurring weekday;
- session start/end;
- session length;
- status/warnings.

Include Unplaced Clients rather than silently omitting them.

Use deterministic Client-name ordering.

### 15.2 Tutor-oriented roster

Group/order assignments by:

1. Tutor name;
2. recurring weekday;
3. session start;
4. Client name.

Include Tutor capacity and enough details for the coordinator to understand each Tutor's assigned Client load.

Provide local export actions for:

- `.md`
- `.pdf`

Use shared report rendering/export infrastructure without adopting another module's report schema.

## 16. Persistence, revisions, and isolation

Tutor Scheduling data belongs to the active Scheduling Session but remains completely module-owned.

Persist enough Tutor-specific revision/provenance information to determine whether a Tutor schedule is stale relative to Tutor Scheduling inputs.

Do not tie Tutor staleness or result validity to revisions in Clinical, Accompanist, or Jury data.

Use versioned schema migrations. Keep `.mpsession` compatibility/versioning and recovery architecture intact.

A Tutor Scheduling result should use a typed Tutor-owned result contract/payload even if its underlying candidate selection was performed by a shared optimizer.

### 16.1 Explicit data-isolation invariant

The following must always hold:

> No Tutor Scheduling domain record has a persisted identity relationship to a domain record owned by another scheduling module.

Consequences:

- matching IDs across modules are coincidental from the application's perspective;
- changing an email/name/availability in Tutor Scheduling never changes another module;
- Tutor Scheduling imports read only into Tutor Scheduling;
- Tutor Scheduling deletes affect only Tutor Scheduling;
- Tutor Scheduling generation reads only Tutor Scheduling source data;
- Tutor Scheduling reports contain only Tutor Scheduling data;
- no hidden synchronization or identity reconciliation is permitted.

Add automated tests around this boundary where practical.

## 17. Time and validation semantics

- Client and Tutor availability recurs weekly throughout the term.
- No date-specific exceptions/holiday model is required initially.
- Availability-grid UI resolution is 30 minutes.
- Automatic candidate start resolution is 15 minutes.
- Client session length is independent of either resolution and defaults to 60 minutes.
- UI clock times use 12-hour AM/PM presentation.
- Use established local-time primitives; do not introduce a new timezone requirement.
- Treat time intervals consistently with shared half-open interval conventions where appropriate.
- Capacity and session length must be positive valid values.
- Invalid imports surface structured validation issues and do not silently invent data.

## 18. Required test coverage

Add module-specific tests at the pure domain/service level before relying on UI tests.

### Data isolation

- identical names/IDs/emails across Tutor and other modules produce independent records;
- Tutor edits/imports/deletes do not mutate Clinical, Jury, or Accompanist data;
- Tutor generation does not read other modules' data;
- Tutor reports contain no data implicitly sourced from other modules.

### Availability editor/model

- Monday-Sunday, 7:00 AM-9:00 PM bounds;
- 30-minute entry grid;
- three-state click cycle;
- right-click to Unavailable;
- click-and-drag painting;
- unmarked -> Unavailable on close/save;
- reuse/generalization does not regress Pianist availability behavior;
- continuous Availability Windows allow 15-minute automatic starts between UI grid boundaries.

### Tags and imports

- uppercase/case-insensitive tag normalization;
- one nonblank tag cell = one tag;
- blank cells create no tag;
- duplicate normalized tags collapse;
- Required AND semantics;
- any Restricted match excludes the automatic candidate;
- each Preferred Client/tag pair contributes one unit;
- Client roster re-import fully replaces Client data, tag states, and assignments without affecting Tutors or any other module.

### Candidate generation and optimizer

- Client session length defaults to 60 minutes;
- client-specific durations are honored;
- both Client and Tutor availability must cover the full session;
- Unavailable on either side excludes a candidate;
- 15-minute automatic starts;
- Tutor capacity is total assigned Clients, independent of overlapping times;
- placement count outranks Preferred matches;
- Preferred matches outrank Tentative-minute minimization;
- Tentative minutes include both Client and Tutor portions with equal per-minute cost;
- mixed Available/Tentative sessions count only Tentative portions;
- earliest start is used only after higher objectives tie;
- global optimization protects hard-to-place Clients;
- identical inputs yield deterministic results;
- shared optimizer has no Tutor/Clinical branching and can be tested using normalized candidates.

### Manual assignment and lifecycle

- manual arbitrary-minute start, such as 8:10 AM;
- calculated end using Client session length;
- manual availability/tag/capacity violations are allowed with warnings;
- manual assignments survive regeneration;
- Manually unplaced Clients remain excluded from regeneration;
- prior automatic assignments are regenerated;
- manual over-capacity assignment leaves zero additional automatic capacity;
- ordinary source edits retain assignments and mark schedule stale;
- deleting an assigned Tutor unplaces only affected Clients;
- deleting a Client removes that Client's assignment;
- Delete All Assignments clears assignments/manual-unplaced/stale-result state without deleting source data.

### Reports/privacy

- Client- and Tutor-oriented report sections;
- deterministic ordering;
- Unplaced/warning visibility;
- local Markdown/PDF export;
- no network dependency.

## 19. Suggested implementation sequence

Keep phases narrow and testable:

1. **Tutor persistence and migrations**: module-owned Tutors, Clients, tag registry, availability associations, Client tag states, assignments, and revision/result state.
2. **Shared optimizer boundary**: if not already extracted for Clinical, define the minimal normalized candidate/capacity/lexicographic optimization core without module-specific types or switches.
3. **Tutors UI/import**: add/edit/delete, capacity, availability grid reuse, tags, Manage Tags, CSV/XLSX mapping/preview.
4. **Client Roster UI/import**: independent Client records, destructive replacement warning, session length default/editing, availability grid reuse, tag states.
5. **Tutor-specific candidate generation/validation**: intersect Tutor and Client availability, session duration, hard tag eligibility, Preferred counts, both-side Tentative-minute calculation.
6. **Tutor optimizer integration**: feed legal Tutor candidates to the shared constrained-assignment core and convert selections into typed Tutor results.
7. **Schedule UI**: generation/regeneration, manual assignments, Manually unplaced, arbitrary-minute manual starts, warnings, diagnostics, stale schedule, Delete All Assignments.
8. **Reports**: Client-oriented and Tutor-oriented Markdown representation plus `.md` and `.pdf` export.
9. **Isolation/regression pass**: prove Tutor data changes do not affect Clinical/Accompanist/Jury data and confirm shared-component extraction does not regress existing modules.

Do not implement the module by copying the Clinical module wholesale and then allowing the copies to diverge. Reuse shared mechanisms where they are genuinely identical, while preserving independent module records and domain adapters.

## 20. Explicit non-goals

Do not add any of the following unless separately approved:

- cross-module Client/Student identity linking;
- synchronization with Clinical, Jury, Accompanist, or Pianist data;
- shared persisted availability records across modules;
- automatic import from another scheduler module;
- simultaneous-session conflict constraints for Clients sharing the same Tutor beyond the stated total Tutor capacity;
- travel time;
- date-specific exceptions/holidays;
- weighted Preferred tags or strong/weak Preferred levels;
- automatic violations of Required/Restricted/Unavailable/capacity rules to increase placements;
- cloud scheduling/storage, telemetry, or analytics;
- a universal persisted scheduling schema;
- module-name branching inside the shared constrained-assignment optimizer.

## 21. Architectural alignment

This specification is subordinate to the repository's canonical product/platform architecture documents for cross-cutting concerns. Tutor Scheduling should follow established Scheduling Session ownership, local-only privacy, versioned migration, shared import/report/availability infrastructure, and module-boundary rules.

Where this document is more specific about Tutor Scheduling domain behavior, it is the implementation specification for this module. Shared code reuse must preserve the core architectural rule: **share mechanisms where justified; keep module-owned data and policy independent.**
