# xlpolish

Finishing touches for Excel files written by Stata's `putexcel` or `export excel`.

`xlpolish` adds what those commands cannot set, so a workbook produced by a
do-file can be handed over ready to use: freeze panes, filters, auto-fit columns,
conditional formatting, dropdown lists, notes, grouping, protected templates,
hidden sheets and print setup, all from one command in your do-file.

## See it in action

<!-- Walkthrough video: edit this file on github.com and drag xlpolish_walkthrough.mp4
     onto this line. GitHub uploads it and replaces this comment with a playable video. -->

## What you need

**Stata 16 or newer. Nothing else.**

`xlpolish` is written in Java and runs on the Java that comes with Stata. Since
Stata 16, Stata installs its own Java runtime alongside the program, so there is
nothing extra to download or configure. In particular, you do **not** need:

- **Microsoft Excel.** `xlpolish` edits the .xlsx file directly, so it works the
  same in Stata for Windows, macOS and Linux, including on servers and remote
  machines without Office.
- **Python, R, or any of their packages.** No `python:` blocks, no `rcall`, and no
  separate installation to keep in sync with Stata.
- **PowerShell, VBA macros or Office automation.** Nothing is scripted through
  Windows or Excel, so it also runs on locked-down work computers where those are
  restricted.
- **Extra Java libraries.** It uses only the standard Java library; the whole
  engine is one small file (`xlpolish.jar`, about 65 KB) installed with the
  command.
- **Administrator rights.** `net install` puts the files in your personal ado
  folder.
- **An internet connection to run it.** Only installing and updating go online.

The workbooks it writes are standard .xlsx files that open in Excel and
LibreOffice like any other.

## Options at a glance

All options can be combined in one call. See `help xlpolish` for details.

| Area | Options |
|---|---|
| Which sheet | `sheet()`, `allsheets` (default: first worksheet) |
| View | `freeze()`, `hidegrid` / `showgrid`, `zoom()`, `tabcolor()`, `hiderows()` / `hidecols()`, `showrows()` / `showcols()` |
| Data | `autofilter`, `filterrange()`, `removefilter`, `nofilterarrow()`, `autofit`, `autofitmax()`, `colwidth()` |
| Conditional formatting | `colorscale()`, `databar()`, `highlight()`, `iconset()`, `top()` / `bottom()`, `average()` |
| Data entry | `dropdown()`, `validate()`, `prompt()`, `unlock()` |
| Notes | `comment()`, `commentauthor()`, `removecomments` |
| Grouping | `grouprows()`, `groupcols()`, `collapse`, `ungroup`, `summaryabove` / `summaryleft` |
| Protection and visibility | `protect`, `password()`, `unprotect`, `hidesheet` / `showsheet` |
| Printing | `landscape` / `portrait`, `fitwidth`, `printtitles()`, `header()` / `footer()` |

## Install

Run this once in Stata:

```stata
net install xlpolish, from("https://raw.githubusercontent.com/fahad-mirza/xlpolish/main/") replace
```

Then `help xlpolish`. To update later, run the same line again and restart Stata.

## Example

```stata
sysuse auto, clear
export excel using "cars.xlsx", firstrow(variables) sheet("Data") replace

xlpolish using "cars.xlsx", freeze(B2) autofilter autofit hidegrid zoom(90)   ///
    colorscale(B2:B75) highlight(C2:C75 >= 30 lightgreen darkgreen)           ///
    comment(B1 "Price in 1978 US dollars") grouprows(54:75)                   ///
    landscape fitwidth printtitles(1)
```

A protected data-entry template, where only the answer cells can be edited and the
choices come from a hidden sheet:

```stata
xlpolish using "template.xlsx", sheet("Form") dropdown(B2:B200 =Lists!A1:A3) ///
    unlock(B2:C200) protect password(team2026)
xlpolish using "template.xlsx", sheet("Lists") hidesheet
```

## Test it

```stata
net get xlpolish, from("https://raw.githubusercontent.com/fahad-mirza/xlpolish/main/")
do xlpolish_test.do
```

It should end with `ALL TESTS PASSED` and leave two workbooks in
`xlpolish_test_output/`:

- **xlpolish_demo.xlsx** has every feature on. Open it in Excel and check: no
  "We found a problem with some content" message. On **Data**: header row and
  `make` column stay put when scrolling; filter arrows on the header; columns sized
  to fit, with header names clear of the arrows; no gridlines; zoom 90%; navy tab;
  price shaded red (cheapest) to green; blue bars on weight; mpg of 30+ green;
  rep78 of 1-2 pink, with empty rep78 cells left white; a Yes/No dropdown in
  column M; red note triangles on the price and mpg headers; a `-` button that
  folds rows 54-75; in print preview, landscape, one page wide, header row on every
  page, and a footer with the file name, page numbers and date. On **Template** (orange tab): typing into column A is refused, while
  B2:D6 can be edited; selecting B2 shows "Choose Yes, No or Unsure" and B2:B6 offer
  those choices; typing 150 into D2 is refused with "Enter a number from 0 to 100."; Review > Unprotect Sheet
  asks for the password `demo`. The **Lists** sheet is hidden (right-click a tab >
  Unhide shows it). On **Scores** (teal tab): two column groups, Ind 1.1-1.3 and
  Ind 2.1-2.3, fold separately with their own -/+ buttons; the slim column E
  between them keeps them apart and has no filter arrow, while every other header
  has one.
- **xlpolish_test.xlsx** is left as the tests finish with it; some settings were
  removed again on purpose.

To test a downloaded copy of this repository instead of an installed one, set
`global XLP_KIT` at the top of `xlpolish_test.do` to the folder; outputs then go
to `test_output/` inside it.

## How it works

An .xlsx file is a zip of XML parts. `xlpolish` edits only the parts it needs (the
target worksheets, plus the workbook, styles, notes and their links when an option
requires them), placing each element where the file format's schema requires.
Every other part is copied through byte for byte. All options are checked before
anything changes, and the result is written to a temporary file in the same folder
and then swapped in, so an error never leaves a half-written workbook.

## Repository layout

| Path | Contents |
|---|---|
| `xlpolish.ado`, `xlpolish.sthlp`, `xlpolish.jar` | the command (installed by `net install`) |
| `xlpolish.pkg`, `stata.toc` | package description for `net install` / `net get` |
| `xlpolish_test.do` | Stata end-to-end test (fetched by `net get`) |
| `java/src/xlpolish/` | Java source (Java 8 compatible) |
| `java/stub/` | compile-time stand-in for Stata's Java API; not shipped in the jar |
| `java/build.sh`, `java/build.bat` | rebuild `xlpolish.jar` (needs a JDK 9+) |
| `tests/` | Java-side test suite (see `tests/README.md`) |
| `CHANGELOG.md` | version history |

## Development

After changing the Java source, rebuild the jar and run both test suites:

```sh
./java/build.sh
./tests/run_tests.sh      # Java side: openpyxl + LibreOffice checks
```

then run `xlpolish_test.do` in Stata and open the demo workbook in Excel.
When publishing a new version, update the version line at the top of
`xlpolish.ado` and `xlpolish.sthlp`, the `Distribution-Date` in `xlpolish.pkg`,
and `CHANGELOG.md`.

## Author

Fahad Mirza ([github.com/fahad-mirza](https://github.com/fahad-mirza)). Questions,
bug reports and suggestions are welcome as
[issues](https://github.com/fahad-mirza/xlpolish/issues); please include the
`xlpolish` command you ran and, if possible, the workbook before it was polished.

