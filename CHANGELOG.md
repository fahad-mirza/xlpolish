# Changelog

## 1.2.0 (2026-10-30)
- New: `summaryabove` / `summarybelow` and `summaryleft` / `summaryright` choose which side of each group its -/+ button and summary row or column sit on (Excel's "Summary rows below detail" and "Summary columns to right of detail" settings). `collapse` marks the matching side.

## 1.1.1 (2026-10-30)
- `grouprows()` / `groupcols()` show a note when two groups touch. Excel stores grouping only as a level per row or column, so touching groups (e.g. C:E and F:H) become one; a gap row or column keeps them separate. Nested groups are not affected.

## 1.1.0 (2026-10-30)
- New: cell notes (`comment()`, `commentauthor()`, `removecomments`); existing notes are kept and merged.
- New: row and column grouping (`grouprows()`, `groupcols()`, `collapse`, `ungroup`).
- New: sheet protection (`protect`, `password()`, `unprotect`) and editable cells on protected sheets (`unlock()`).
- New: hide and show sheets (`hidesheet`, `showsheet`), keeping at least one sheet visible and moving the active tab if needed.
- `dropdown(range =Sheet!A1:A5)` now makes the cell reference fixed (`$A$1:$A$5`) itself, so no `$` has to be typed in Stata.
- Semicolons inside double quotes no longer split rules, so note and highlight text may contain them.
- The temporary option file is erased as soon as the Java step finishes (it may hold a password).
- Repository laid out for `net install` from GitHub; the test do-file also runs on an installed copy.

## 1.0.3
- `autofit` leaves room for filter dropdown arrows in header cells.

## 1.0.2
- `highlight()` never shades empty cells (Excel's own comparison rules treat them as 0).

## 1.0.1
- Removed a stray "file not found" note shown on every run.

## 1.0.0
- First version: freeze panes, gridlines, zoom, tab colour, filters, auto-fit and set widths,
  colour scales, data bars, highlight rules, dropdowns, orientation, fit to width, print titles.
