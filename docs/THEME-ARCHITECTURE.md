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

The built-in theme has exactly ten base colors. They are the only values a future
palette file supplies:

```text
primary
accent
background
surface
text
textMuted
border
success
warning
error
```

All other visual roles are semantic aliases or derived values. They must not
become additional palette entries:

| Semantic role | Base color or derivation |
|---|---|
| Header, links, focus ring, information, jury-required-badge | `primary` |
| Secondary actions | `accent` |
| Elevated surface, panel, dialog, tooltip, neutral tag | `surface` |
| Inverse text | `surface` |
| Muted and disabled text | Readable derived mix of `textMuted` and `text` |
| Border strong | Derived from `border` and `textMuted` |
| Available, ready, required constraints/tags | `success` |
| Tentative, warning, stale, preferred tag | `warning` |
| Error, conflict, blocked, restricted tag | `error` |
| Unavailable | Derived from `textMuted`, `border`, and `surface`; it is not an error |
| Selection and status backgrounds | Derived tints of the relevant base color and `surface` |
| Hover and active states, shadows, overlays | Derived color mixes or opacity |

Use the same alias for the same meaning in every module. Status must not rely
on color alone.

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
  primary: string;
  accent: string;
  background: string;
  surface: string;
  text: string;
  textMuted: string;
  border: string;
  success: string;
  warning: string;
  error: string;
}
```

The palette object intentionally contains only these ten base values. Runtime
CSS variables such as hover colors, status aliases, and surface tints are
derived from them.

---

# Palette PNG Format

Theme palettes are authored as PNG files.

The PNG is the source artwork.

The application derives theme colors from fixed pixel locations.

The PNG is not displayed to users.

---

## Cell Layout

Each palette entry occupies an 8x8 pixel square.

A theme PNG contains a fixed grid and must be 40x16 pixels. The parser accepts
non-interlaced, 8-bit RGB or RGBA PNGs. RGBA palette samples must be fully
opaque.

The palette PNG contains exactly ten cells, each an 8x8 solid-color square:

```text
Row 1: Primary | Accent | Background | Surface | Text
Row 2: Muted Text | Border | Success | Warning | Error
```

Coordinates are zero-based; each entry is sampled from the center pixel of its
8x8 cell. The parser verifies the PNG signature, chunk boundaries and CRCs,
image dimensions, supported encoding, decompressed image data, and that all ten
sampled cells contain opaque colors. Optional true-color PLTE chunks are
accepted; transparency-key (`tRNS`) chunks are not.

In a future theme-loading flow, changing the PNG changes the derived base
palette values.

---

# Theme Loading

Theme loading must occur entirely locally.

No network access.

No cloud service.

No telemetry.

Themes are ordinary local application assets.

The theme subsystem defines `ThemePalette` in
`webapp/frontend/src/theme/ThemePalette.ts`. It contains exactly the ten base
colors listed above. The default palette PNG is bundled at
`webapp/frontend/src/theme/default-theme.png` and reproduces the built-in
palette. The model's built-in values remain the fallback if the asset cannot be
loaded or validated.

The PNG parser in `webapp/frontend/src/theme/pngPalette.ts` checks the PNG
signature, chunk boundaries and CRCs, image dimensions, supported encoding,
decompressed image data, and all ten opaque center-pixel samples. It maps
zero-based pixel coordinates to palette values as follows:

| Palette value | Sample coordinate `(x, y)` |
|---|---:|
| `primary` | `(4, 4)` |
| `accent` | `(12, 4)` |
| `background` | `(20, 4)` |
| `surface` | `(28, 4)` |
| `text` | `(36, 4)` |
| `textMuted` | `(4, 12)` |
| `border` | `(12, 12)` |
| `success` | `(20, 12)` |
| `warning` | `(28, 12)` |
| `error` | `(36, 12)` |

`themeValidation.ts` validates the exact ten-color palette structure and
provides contrast checks for body and muted text, success/warning/error
messages, focus indicators, and control borders. Contrast issues are structured
validation results; a palette that fails is not applied. The loader applies
only the ten base CSS variables. Derived and semantic variables remain CSS
definitions and update automatically when base variables change.

The PNG parser only decodes and validates a palette image; it does not select a
theme, expose file selection, or change the active application theme. A load or
validation failure returns the same built-in palette with a structured fallback
warning containing the failure reason and, when available, its error code.

PNG decoding currently supports non-interlaced 8-bit RGB and RGBA data. Other
valid PNG encodings are rejected with a clear validation error rather than
being sampled incorrectly.

---

## Startup and Theme Selection

The dashboard provides a fixed lower-right Theme button that opens a native
modal dialog with Default Theme, Dark Theme, and Custom Theme options. Each
choice shows a palette preview strip and indicates the active theme. The native
dialog provides keyboard focus handling and Escape-to-close behavior.

The bundled default and dark theme PNGs are loaded locally through the same
parser and accessibility validator. The dark theme uses these ten base values:

| Palette value | Dark theme |
|---|---|
| `primary` | `#80BFFF` |
| `accent` | `#62A3E8` |
| `background` | `#111820` |
| `surface` | `#1B2633` |
| `text` | `#F2F4F7` |
| `textMuted` | `#B4BFCC` |
| `border` | `#78889B` |
| `success` | `#67D391` |
| `warning` | `#FFD166` |
| `error` | `#FF7B7B` |

Custom themes are selected from local PNG files, parsed in the frontend, and
applied immediately only after structure and accessibility checks pass.
Invalid files display validation errors without changing the active theme.

The platform abstraction persists a versioned active-theme preference in
local application storage. Built-in themes persist their identifier; custom
themes persist only the ten validated palette values needed to restore them.
No theme file or palette is uploaded or synchronized. Startup restores the
saved choice locally. Invalid preferences, failed decoding, or inaccessible
saved palettes fall back to the built-in default with a structured warning.
The active palette sets the native `color-scheme` attribute from its background
luminance so built-in controls adapt to dark surfaces. Semantic roles and
derived CSS mappings do not change.

Theme initialization failure never blocks application startup.

---

# CSS Integration

All visual styling should ultimately flow through theme variables.

Preferred:

```css
--theme-primary
--theme-background
--theme-success
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

The current validator checks 4.5:1 contrast for normal-size text and 3:1 for
focus indicators and borders. It checks warning and error foregrounds against
their derived soft backgrounds as well as surfaces. Both built-in palettes
must pass these checks. Custom palettes that fail are rejected before
application; invalid saved selections produce a structured fallback warning.

---

# Future Theme Expansion

The dashboard selector currently supports Default, Dark, and one active Custom
palette. Future work may add bundled themes and richer theme management while
continuing to use the same ten-color contract, local validation, and storage
boundary. Editing palettes in-app, sharing themes, and downloading theme
galleries are not implemented.

---

# Future Theme Files

Recommended location:

```text
frontend/src/themes/
```

Examples of future bundled assets include forest or other institution-approved
palettes. The existing selector and parser can load bundled themes without
changing module-owned semantics.

---

# Out of Scope

This architecture does not define:

- in-app palette editing
- cloud theme sharing
- online theme galleries or external theme downloads

Those are separate features. Theme switching between the bundled themes and a
local custom PNG is already implemented.

---

# Migration Strategy

1. Audit all hardcoded colors.
2. Create semantic theme tokens.
3. Replace hardcoded colors with tokens.
4. Create default theme matching the current appearance.
5. Implement PNG palette loading.
6. Add theme switching later.

Preserve current visual appearance until theme switching is intentionally introduced.