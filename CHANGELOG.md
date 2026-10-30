# Changelog

## 1.3.1 (2026-10-30)
- Conditional formats from each call now go above the rules already in the sheet, as when adding rules in Excel, so a later highlight shows over an earlier colour scale or data bar. Within one call, highlight(), top(), bottom() and average() go above colorscale(), databar() and iconset().
- Demo workbook: new Scores sheet showing column groups side by side with a spacer column that has no filter arrow.

## 1.3.0 (2026-10-30)
- New: `nofilterarrow()` hides the filter arrow on chosen columns (e.g. spacer columns between column groups).
- New: `validate()` restricts cells to whole numbers, decimals, dates or text lengths within limits, with a plain default error message or your own.
- New: `prompt()` shows a message when a cell is selected; it joins any dropdown or rule on the same cells.
- New: `iconset()` (traffic lights, arrows, flags, ratings...), `top()` / `bottom()` (N or N%), and `average()` (above or below) conditional formats.
- New: `header()` / `footer()` with [page], [pages], [date], [time], [file] and [sheet] placeholders and left | centre | right sections.
- New: `hiderows()`, `hidecols()`, `showrows()`, `showcols()`.
- `dropdown()` on exactly the same range as an existing dropdown now replaces it instead of adding a second rule; a range that only partly overlaps another rule is refused with an explanation.

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
