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

The theme-loading service owns the `ThemePalette` type, the built-in default
palette, exact ten-color/hex validation, and generation/application of the ten
base CSS variables. Derived and semantic variables remain CSS definitions and
update automatically when base variables change.

The PNG parser only decodes and validates a palette image; it does not select a
theme, expose file selection, or change the active application theme. Loading
without a candidate uses the existing built-in palette. A candidate load or
validation failure returns that same palette with a structured fallback warning
so the caller can log the failure without blocking startup.

PNG decoding currently supports non-interlaced 8-bit RGB and RGBA data. Other
valid PNG encodings are rejected with a clear validation error rather than
being sampled incorrectly.

---

## Startup

The default theme loads automatically at startup.

If loading fails:

- fall back to the built-in default theme
- return a structured warning for the caller to log
- continue application startup

Theme failure must never block startup.

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