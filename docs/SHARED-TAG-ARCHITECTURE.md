# Shared Tag Architecture

## Status

Canonical and authoritative shared infrastructure architecture.

Tags are a shared infrastructure concept within Music Program Scheduler.

The application provides shared tag identity and infrastructure that may be
used by multiple modules. Each module owns its own catalog and tag data. This
document supersedes conflicting tag-related statements elsewhere in the
repository.

This follows the established architectural principle of sharing proven
infrastructure while keeping business rules, workflows, optimization logic,
and domain meaning inside individual modules.

---

# Purpose

Tags provide a flexible classification mechanism that allows modules to
represent:

- Categories
- Specializations
- Capabilities
- Preferences
- Restrictions
- Attributes
- Groupings

without requiring new schema fields whenever users need additional
classifications.

Examples:

Clinical Placements:

- Mental Health
- Pediatrics
- Virtual
- Advanced

Performance Juries:

- Honors
- Graduate
- Composition

Future Modules:

- User-defined classifications

Tags are intended to be:

- User-defined
- Extensible
- Searchable
- Reportable
- Importable
- Exportable

Tags are not scheduling rules.

---

# Architectural Principles

## Shared Infrastructure

The shared tag infrastructure owns:

- Tag identity
- Tag storage
- Tag normalization
- Tag assignment mechanics
- Tag import parsing
- Tag export formatting
- Tag validation primitives
- Tag filtering support
- Tag reporting support

This infrastructure is shared across all modules.

The infrastructure supplies mechanisms and primitives, not a global tag
catalog, taxonomy, or business policy. Module-qualified identity is mandatory:
every tag is identified by `(module_id, tag_uuid)`. Catalog operations and
assignments must remain scoped to that module identity.

The goal is to provide consistent shared mechanisms while allowing
module-owned workflows and presentation.

---

## Module Ownership

Modules own:

- Tag catalogs
- Tag meanings
- Tag assignment semantics
- Validation policies
- Optimization behavior
- User workflows
- Reporting behavior

The same tag name may exist in multiple modules.

Those tags are unrelated.

Example:

```text
Clinical Placements
Advanced

Performance Juries
Advanced
```

These are distinct tags with distinct identities.

The application must not assume any relationship between them.

---

# Core Concepts

## Tag

A Tag is a reusable classification value.

### Fields

| Field | Description |
|---------|---------|
| Tag UUID | Internal identifier |
| Module ID | Owning module |
| Display Name | User-facing label |
| Normalized Name | Canonical comparison value |
| Created Timestamp | Creation time |
| Modified Timestamp | Last modification time |

Example:

```text
Display Name:
Mental Health

Normalized Name:
mental health
```

---

## Tag Identity

Tags are uniquely identified by:

```text
(module_id, tag_uuid)
```

Tag names are presentation values only.

The following is valid:

```text
Clinical:
Advanced

Jury:
Advanced
```

because the tags belong to different modules.

Tag names must only be unique within their owning module.

---

## Tag Catalog

Every module owns a separate tag catalog.

Examples:

### Clinical Placements

```text
Mental Health
Pediatrics
Virtual
Advanced
School-Based
Geriatrics
```

### Performance Juries

```text
Honors
Graduate
Composition
```

### Future Module

```text
Any module-specific values
```

Tag catalogs never cross module boundaries.

Operations in one module must not affect tags in another module.

---

# Normalization Rules

Tag comparisons are case-insensitive.

Before comparison:

1. Trim leading whitespace.
2. Trim trailing whitespace.
3. Collapse repeated internal whitespace.
4. Apply case-insensitive normalization.

The resulting normalized value is for comparison and uniqueness. The display
name remains a presentation value; any module-specific display formatting
must not change tag identity.

Examples:

```text
Mental Health
mental health
MENTAL HEALTH
```

all resolve to the same normalized value.

Example:

```text
Mental    Health
```

becomes:

```text
Mental Health
```

---

# Tag Assignment

Tags may be assigned to module-owned entities.

Examples:

### Clinical Site

```text
Mental Health
Virtual
Advanced
```

### Clinical Student

```text
Required: Mental Health
Preferred: Virtual
Restricted: Pediatrics
```

The shared infrastructure stores assignments.

The owning module determines what those assignments mean.

---

# Assignment States

The shared infrastructure supports assignment states.

The infrastructure stores the state value.

The module defines the meaning.

Example:

Clinical Placements:

```text
Required
Preferred
Restricted
```

Interpretation:

- Required = site must contain tag
- Preferred = site should be favored when possible
- Restricted = site must not contain tag

These meanings belong to the Clinical Placement module.

Another module may use:

```text
Include
Exclude
```

or

```text
Required
Optional
```

or no assignment states at all.

The shared infrastructure must not impose meaning.

---

# Import Format

## Design Principle

Tags must never be inferred from unmapped spreadsheet columns.

The following behavior is prohibited:

```text
Any unmapped column becomes a tag.
```

This behavior is difficult to explain, easy to misuse, and produces
unpredictable imports.

Imports must use explicit tag mappings.

Only source columns explicitly mapped to tag fields defined by the importing
module are parsed as tags. An unmapped column is never implicitly a tag
column. The importing module chooses which tag fields it supports and how
they relate to its entities and assignment states.

---

## Standard Tag Column

Preferred format:

| Site Name | Tags |
|------------|------------|
| Community Clinic | Mental Health; Virtual; Advanced |

Tags are separated by semicolons.

### Parsing Rules

- Split on semicolon (`;`)
- Trim whitespace
- Ignore empty values
- Remove duplicates
- Apply normalization
- Match existing tags case-insensitively
- Create missing tags within the current module

Example:

```text
Mental Health; Virtual ; Advanced;
```

becomes:

```text
Mental Health
Virtual
Advanced
```

---

## Multiple Tag Columns

Modules may define multiple tag-related columns.

Example:

| Student | Required Tags | Preferred Tags | Restricted Tags |
|----------|----------|----------|----------|
| Student A | Mental Health | Virtual | Pediatrics |

Values remain semicolon-separated.

Example:

```text
Required Tags:
Mental Health; School-Based

Preferred Tags:
Virtual

Restricted Tags:
Pediatrics
```

The shared infrastructure parses the values.

The module interprets them.

---

## Module Isolation During Import

Imports operate only within the current module.

Example:

Clinical Placement import:

```text
Mental Health
Virtual
```

creates or matches tags only in the Clinical Placement catalog.

The import must not:

- Create Jury tags
- Modify Jury tags
- Match against Jury tags
- Reference Jury tags

Module boundaries remain intact.

---

# Import Preview

Tag imports participate in standard import preview workflows.

Preview should identify:

- Existing tags matched
- New tags created
- Duplicate values removed
- Invalid values
- Empty values ignored

Example:

Source:

```text
Mental Health; Mental Health; Virtual
```

Preview:

```text
Matched:
- Mental Health
- Virtual

Duplicates Removed:
- Mental Health
```

Users should be able to review newly created tags before applying the import.

---

# Export Format

Exports should use the same semicolon-separated representation.

Example:

```text
Mental Health; Virtual; Advanced
```

Exports include only tags belonging to the exporting module.

This preserves round-trip import/export behavior.

---

# Editing

Users may:

- Create tags
- Rename tags
- Delete tags
- Assign tags
- Remove tag assignments

Renaming a tag updates references within the owning module.

Tag UUIDs remain stable.

Tag operations never affect other modules.

---

# Validation

## Required Rules

Tag names:

- Must not be blank
- Must not exceed configured maximum length
- Must be unique within the owning module
- Must be case-insensitively unique within the owning module

### Invalid

```text
""
```

### Invalid

```text
Mental Health
mental health
```

within the same module.

### Valid

Clinical:

```text
Advanced
```

Jury:

```text
Advanced
```

because they belong to different module catalogs.

---

# Reporting

Tags may participate in:

- Filtering
- Sorting
- Grouping
- Report generation

The shared reporting infrastructure may expose tag fields.

Each module determines:

- Which tags are reportable
- Which tags are filterable
- Which tags are groupable
- How tags appear in reports

Cross-module tag reporting is not supported.

---

# Persistence and Migration

Shared persistence must preserve the `(module_id, tag_uuid)` identity and
module-scoped catalog uniqueness. Tag assignment mechanics may be shared,
while entity relationships and any assignment-state values remain scoped to
the owning module. No tag assignment creates a relationship between modules.

Existing module-specific tag tables and references must be brought into
conformance through the application's ordered, forward-only migration
registry. Migration must preserve each module's tag UUIDs, catalog entries,
assignments, and module-specific states, and must support staged `.mpsession`
restore. It must not merge tags by display name across modules. See
[ARCHITECTURE.md](ARCHITECTURE.md) for current schema and migration planning.

---

# Clinical Placement Usage

## Clinical Sites

Import example:

```text
Site Name
Tags

Community Clinic
Mental Health; Virtual; Advanced
```

Site tags describe site characteristics.

---

## Clinical Students

Import example:

```text
Student Name
Required Tags
Preferred Tags
Restricted Tags
```

Example values:

```text
Required Tags:
Mental Health

Preferred Tags:
Virtual

Restricted Tags:
Pediatrics
```

Student tag assignments describe placement requirements and preferences.

---

## Clinical Scheduling Semantics

Clinical Placement interprets tag states as follows.

### Required

A placement site must contain the tag.

### Restricted

A placement site must not contain the tag.

### Preferred

A placement site containing the tag receives a higher preference score.

These behaviors belong exclusively to the Clinical Placement module.

They are not part of the shared tag infrastructure.

`Required`, `Preferred`, and `Restricted` are Clinical assignment-state
semantics only. They are not global states, and another module may define
different states or none. Clinical tag identity remains
`(clinical-placements, tag_uuid)`; an identically named tag in another module
is unrelated.

---

# Future Modules

Future modules may reuse the shared tag infrastructure.

They may define:

- Their own tag catalogs
- Their own assignment states
- Their own validation rules
- Their own reporting behavior
- Their own optimization behavior

The shared infrastructure provides consistency.

The module provides meaning.

---

# Explicitly Out of Scope

The shared tag infrastructure does not provide:

- Universal optimization behavior
- Universal scoring
- Universal validation policies
- Cross-module tag catalogs
- Cross-module tag inheritance
- Cross-module tag references
- Automatic workflow behavior based on tag names
- Automatic scheduling rules

A tag is a classification value.

A tag is not a scheduling rule.

Modules remain responsible for determining how tags affect their own
validation, optimization, workflow, and reporting behavior.

---

# Architectural Summary

Shared Infrastructure:

- Tag model
- Tag identity
- Tag storage
- Tag assignments
- Tag parsing
- Tag validation primitives
- Tag import/export support
- Tag filtering support
- Tag reporting support

Module-Owned:

- Tag catalogs
- Tag meanings
- Tag states
- Validation policies
- Optimization behavior
- Workflow behavior
- Reporting semantics

This mirrors the established Music Program Scheduler pattern of shared
infrastructure with module-owned business logic.