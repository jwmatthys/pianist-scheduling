# THEME-ARCHITECTURE.md

# Scheduling Maestro Theme Architecture

**Status:** Approved design. Implementation may proceed incrementally.

## Purpose

Scheduling Maestro uses a shared theming system that provides visual consistency across all modules while allowing rapid experimentation with alternate color palettes.

The theme system is shared infrastructure.

Individual scheduling modules:

- Accompanist Scheduling
- Lesson Scheduling
- Tutor Scheduling
- Performance Juries
- Clinical Placements

must consume the same theme tokens and must not define independent color systems.

This document defines:

- theme ownership
- semantic color roles
- palette-file format
- theme-loading behavior
- future theme switching

It does not define specific colors.

---

# Design Principles

## Semantic, Not Color-Based

UI code must never depend on specific colors.

Avoid:

```css
--blue-500
--green-300
--red-700
```

Use:

```css
--color-primary
--color-success
--color-warning
--color-error
```

Components consume meaning.

Themes provide appearance.

---

## Shared Infrastructure

Theme infrastructure is shared.

Theme values are not module-specific.

The following concepts must be shared:

- color tokens
- theme loading
- palette parsing
- CSS variable generation

Modules consume the resulting theme.

They do not define colors.

---

## Visual Consistency

The same semantic meaning should use the same token throughout the application.

Examples:

Success:
- readiness passed
- import succeeded
- generation succeeded

Warning:
- stale result
- tentative availability
- validation warning

Error:
- invalid data
- readiness blocker
- failed operation

A module should not invent a different color for the same semantic meaning.

---

# Theme Tokens

## Application Surfaces

```text
background
surface
surfaceElevated
panel
header
sidebar
dialog
tooltip
```

---

## Text

```text
textPrimary
textSecondary
textMuted
textDisabled
textInverse
```

---

## Actions

```text
primary
primaryHover
primaryActive

secondary
secondaryHover
secondaryActive

link
linkHover
```

---

## Status

```text
success
warning
error
info
```

---

## Scheduling States

Shared scheduling concepts use dedicated tokens.

```text
available
tentative
unavailable
```

These are shared infrastructure concepts and may appear in multiple modules.

---

## Tag States

```text
tagRequired
tagPreferred
tagRestricted
tagNeutral
```

---

## Validation

```text
conflict
stale
ready
blocked
```

---

## Borders and Focus

```text
border
borderStrong
focusRing
selection
```

---

# Theme Storage

Themes are represented internally as:

```ts
ThemePalette
```

or equivalent.

Conceptually:

```ts
interface ThemePalette {
  background: string;
  surface: string;
  panel: string;

  textPrimary: string;
  textSecondary: string;
  textMuted: string;

  primary: string;
  primaryHover: string;
  primaryActive: string;

  success: string;
  warning: string;
  error: string;
  info: string;

  available: string;
  tentative: string;
  unavailable: string;

  tagRequired: string;
  tagPreferred: string;
  tagRestricted: string;
  tagNeutral: string;

  conflict: string;
  stale: string;
  ready: string;
  blocked: string;

  border: string;
  borderStrong: string;
  focusRing: string;
  selection: string;
}
```

Actual implementation may contain additional fields.

---

# Palette PNG Format

Theme palettes are authored as PNG files.

The PNG is the source artwork.

The application derives theme colors from fixed pixel locations.

The PNG is not displayed to users.

---

## Cell Layout

Each palette entry occupies an 8x8 pixel square.

A theme PNG contains a fixed grid.

Example:

```text
Row 1
Primary
Primary Hover
Primary Active
Secondary
Secondary Hover
Secondary Active

Row 2
Background
Surface
Elevated Surface
Panel
Header
Sidebar

Row 3
Primary Text
Secondary Text
Muted Text
Disabled Text
Inverse Text

Row 4
Success
Warning
Error
Info

Row 5
Available
Tentative
Unavailable

Row 6
Required Tag
Preferred Tag
Restricted Tag
Neutral Tag

Row 7
Conflict
Stale
Ready
Blocked

Row 8
Border
Border Strong
Focus Ring
Selection
```

The exact coordinate mapping must be documented in code.

Changing the PNG changes the theme.

---

# Theme Loading

Theme loading must occur entirely locally.

No network access.

No cloud service.

No telemetry.

Themes are ordinary local application assets.

---

## Startup

The default theme loads automatically at startup.

If loading fails:

- fall back to the built-in default theme
- continue application startup
- log a structured warning

Theme failure must never block startup.

---

# CSS Integration

All visual styling should ultimately flow through theme variables.

Preferred:

```css
--color-primary
--color-background
--color-success
```

Components should consume variables.

They should not contain hardcoded colors.

---

# Accessibility

Every theme must maintain:

- readable text
- visible focus indicators
- sufficient contrast

Theme loading does not bypass accessibility requirements.

A palette that produces unreadable UI is invalid.

Validation may be added in the future.

---

# Future Theme Switching

Future versions may allow:

```text
Settings
  Theme
    Default
    Dark
    Forest
    Custom
```

This architecture supports theme switching but does not require it.

Theme switching is a separate feature.

---

# Future Theme Files

Recommended location:

```text
frontend/src/themes/
```

Examples:

```text
default-theme.png
forest-theme.png
dark-theme.png
```

The active theme may be selected from settings in a future milestone.

---

# Out of Scope

This architecture does not define:

- specific color values
- dark mode behavior
- theme switching UI
- cloud theme sharing
- user-created theme editing tools

Those are separate features.

---

# Migration Strategy

1. Audit all hardcoded colors.
2. Create semantic theme tokens.
3. Replace hardcoded colors with tokens.
4. Create default theme matching the current appearance.
5. Implement PNG palette loading.
6. Add theme switching later.

Preserve current visual appearance until theme switching is intentionally introduced.