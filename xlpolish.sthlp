{smcl}
{* *! version 1.3.1  30oct2026}{...}
{title:Title}

{phang}
{bf:xlpolish} {hline 2} Finishing touches for Excel files written by {helpb putexcel} or {helpb export excel}


{title:Syntax}

{p 8 16 2}
{cmd:xlpolish} {cmd:using} {it:filename}{cmd:,} [{it:options}]

{synoptset 30 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:Which sheet}
{synopt:{opt sheet(name)}}worksheet to change (not case-sensitive); default is the first worksheet{p_end}
{synopt:{opt allsheets}}apply to every worksheet (chart sheets are skipped){p_end}

{syntab:View}
{synopt:{opt freeze(cell)}}freeze rows above and columns left of {it:cell}; {cmd:freeze(A1)} removes{p_end}
{synopt:{opt hidegrid} | {opt showgrid}}hide or show the cell gridlines{p_end}
{synopt:{opt zoom(#)}}zoom level in percent, 10-400{p_end}
{synopt:{opt tabcolor(color)}}colour of the sheet tab; {cmd:tabcolor(none)} removes{p_end}
{synopt:{opt hiderows(rows)} | {opt showrows(rows)}}hide or show rows, e.g. {cmd:hiderows(5:7)}{p_end}
{synopt:{opt hidecols(cols)} | {opt showcols(cols)}}hide or show columns, e.g. {cmd:hidecols(E H:J)}{p_end}

{syntab:Data}
{synopt:{opt autofilter}}filter dropdowns on the used range; its first row is the header{p_end}
{synopt:{opt filterrange(range)}}filter dropdowns on {it:range}, e.g. {cmd:A3:H164}{p_end}
{synopt:{opt removefilter}}remove the filter{p_end}
{synopt:{opt nofilterarrow(cols)}}no filter arrow on these columns of the filter, e.g. spacer columns{p_end}
{synopt:{opt autofit}}size each used column to its longest value{p_end}
{synopt:{opt autofitmax(#)}}widest column {opt autofit} may set; default 60{p_end}
{synopt:{opt colwidth(spec)}}set widths, e.g. {cmd:colwidth(A=30 B:D=12)}{p_end}

{syntab:Conditional formatting and dropdowns}
{synopt:{opt colorscale(range [colors])}}shade cells by value with 2 or 3 colours{p_end}
{synopt:{opt databar(range [color])}}in-cell bars proportional to value{p_end}
{synopt:{opt highlight(range op value fill [text])}}fill cells that meet a condition{p_end}
{synopt:{opt iconset(range style)}}traffic lights, arrows, flags or ratings by value{p_end}
{synopt:{opt top(range N fill [text])}}fill the N (or N%) highest values{p_end}
{synopt:{opt bottom(range N fill [text])}}fill the N (or N%) lowest values{p_end}
{synopt:{opt average(range above|below fill [text])}}fill values above or below the average{p_end}

{syntab:Data entry}
{synopt:{opt dropdown(range choices)}}in-cell dropdown list{p_end}
{synopt:{opt validate(range type rule)}}allow only numbers, dates or text lengths within limits{p_end}
{synopt:{opt prompt(range "text")}}message shown when one of the cells is selected{p_end}

{syntab:Notes}
{synopt:{opt comment(cell "text")}}add or replace a note (hover pop-up) on a cell; {cmd:""} deletes it{p_end}
{synopt:{opt commentauthor(name)}}author shown with new notes; default is your user name{p_end}
{synopt:{opt removecomments}}delete every note on the sheet{p_end}

{syntab:Grouping}
{synopt:{opt grouprows(rows)}}group rows so they can be collapsed, e.g. {cmd:grouprows(5:20 25:40)}{p_end}
{synopt:{opt groupcols(cols)}}group columns, e.g. {cmd:groupcols(C:F)}{p_end}
{synopt:{opt collapse}}start the new groups collapsed{p_end}
{synopt:{opt ungroup}}remove all row and column grouping{p_end}
{synopt:{opt summaryabove} | {opt summarybelow}}put each row group's button and summary row above or below it (below is Excel's default){p_end}
{synopt:{opt summaryleft} | {opt summaryright}}put each column group's button and summary column left or right of it (right is Excel's default){p_end}

{syntab:Protection and visibility}
{synopt:{opt protect}}protect the sheet against editing{p_end}
{synopt:{opt password(text)}}password needed to unprotect it; requires {opt protect}{p_end}
{synopt:{opt unprotect}}remove protection{p_end}
{synopt:{opt unlock(ranges)}}keep these cells editable when the sheet is protected{p_end}
{synopt:{opt hidesheet} | {opt showsheet}}hide or show the sheet's tab{p_end}

{syntab:Printing}
{synopt:{opt landscape} | {opt portrait}}page orientation{p_end}
{synopt:{opt fitwidth}}fit all columns on one printed page width{p_end}
{synopt:{opt printtitles(rows [cols])}}repeat rows and/or columns on every printed page, e.g. {cmd:printtitles(1)} or {cmd:printtitles(1:2 A)}{p_end}
{synopt:{opt header("text")} | {opt footer("text")}}page header or footer, with page numbers, date, file or sheet name{p_end}
{synoptline}


{title:Description}

{pstd}
{cmd:xlpolish} adds the Excel features that {helpb putexcel} and {helpb export excel}
cannot set, so a workbook produced by a do-file can be handed over ready to use:
filter dropdowns, auto-fit column widths, conditional formatting, dropdown lists,
print settings, tab colours and zoom. It also adds cell notes, row and column
grouping, sheet protection with editable cells, and hidden sheets, and sets
freeze panes and gridlines, which {cmd:putexcel} only offers from Stata 18 onwards.

{pstd}
Run it after the workbook has been written. All options in one call are applied
together and the file is saved once. Options are checked before anything is
changed: if any option is invalid, the workbook is left exactly as it was.

{pstd}
{cmd:xlpolish} runs on Stata's built-in Java. It needs no Excel, Python, or
PowerShell, and only the parts of the file it changes are rewritten; data,
formulas, formatting, images and charts are copied through untouched. The
workbook must be closed in Excel while {cmd:xlpolish} runs.


{title:Options}

{dlgtab:Which sheet}

{phang}
{opt sheet(name)} names the worksheet to change. Matching is not
case-sensitive. The default is the first worksheet in the workbook.

{phang}
{opt allsheets} applies the options to every worksheet. Ranges you give (for
example in {opt highlight()}) are used on every sheet.

{dlgtab:View}

{phang}
{opt freeze(cell)} keeps everything above and to the left of {it:cell} visible
while scrolling: {cmd:freeze(A2)} freezes the header row, {cmd:freeze(B2)} the
header row and column A, {cmd:freeze(A1)} removes any freeze.

{phang}
{opt hidegrid} hides the cell gridlines; {opt showgrid} shows them again.

{phang}
{opt zoom(#)} sets the zoom level, from 10 to 400 percent.

{phang}
{opt hiderows(rows)} and {opt hidecols(cols)} hide rows (such as {cmd:5} or
{cmd:5:20}) and columns (such as {cmd:E} or {cmd:H:J}) without deleting them,
e.g. helper columns readers do not need. {opt showrows()} and {opt showcols()}
show them again. Hidden rows and columns are still read by {helpb import excel}.

{phang}
{opt tabcolor(color)} colours the sheet tab. {it:color} may be any colour name
accepted by {helpb putexcel} (for example {cmd:navy}, {cmd:lightgreen}), a hex code
({cmd:1F4E79}), or an RGB triple in quotes ({cmd:"31 78 121"}).
{cmd:tabcolor(none)} removes the colour.

{dlgtab:Data}

{phang}
{opt autofilter} adds filter dropdowns to the header row of the used range (the
smallest rectangle holding every non-empty cell); the first row of that range is
treated as the header. If the table has a title above the header, use
{opt filterrange()} instead. Not allowed on sheets that contain an Excel table,
which has its own filter.

{phang}
{opt filterrange(range)} adds filter dropdowns to {it:range}, whose first row is
the header, e.g. {cmd:filterrange(A3:H164)}.

{phang}
{opt removefilter} removes the filter from the sheet.

{phang}
{opt nofilterarrow(cols)} hides the filter dropdown arrow on the listed columns,
such as {cmd:E} or {cmd:E I}. A sheet has one filter, covering one continuous
block of columns, so spacer columns between column groups have to be inside it;
this keeps them free of arrows. The columns must lie within the filter, added
in the same call ({opt autofilter} or {opt filterrange()}) or already there.
Adding {opt autofilter} again later recreates the filter with all arrows, so
give {opt nofilterarrow()} in that call too.

{phang}
{opt autofit} sets each used column's width from its longest displayed value.
Excel files do not store a true auto-fit, so widths are estimated from the
number of characters, which gives a close approximation for standard fonts.
Cells inside merged ranges that span several columns are ignored. When the
sheet has a filter (added in the same call or already there), header cells get
extra room for the dropdown arrow so the column names are not hidden behind it. Numbers are
measured as they would appear in Excel's General format, so columns with a
number format that adds characters (such as thousands separators) may need a
little extra room; add it with {opt colwidth()}.

{phang}
{opt autofitmax(#)} caps the widths set by {opt autofit}; the default is 60
characters. Long text above the cap is cut off on screen but not in the cell.

{phang}
{opt colwidth(spec)} sets column widths in characters, 0-255. {it:spec} is a
list such as {cmd:A=30 B:D=12 H=8}. {opt colwidth()} is applied after
{opt autofit}, so it overrides it for the columns it names.

{dlgtab:Conditional formatting and dropdowns}

{pstd}
Ranges may be single cells ({cmd:B2}), rectangles ({cmd:B2:H164}), whole columns
({cmd:B:D}) or whole rows ({cmd:2:5}). Each of the four options below accepts
several rules separated by semicolons, e.g.
{cmd:highlight(D2:D164 == 5 lightgreen; D2:D164 == 1 FFC7CE)}. Rules are added to
any that are already in the sheet, so running the same command twice adds the
rule twice.

{phang}
{opt colorscale(range [low [middle] high])} shades each cell by its value
relative to the others in {it:range}. With no colours, the Excel default
red-yellow-green scale is used (lowest red). With two colours, the scale runs
from the lowest to the highest value; with three, the middle colour marks the
median.

{phang}
{opt databar(range [color])} draws a bar in each cell proportional to its value.
The default colour is Excel's blue.

{phang}
{opt highlight(range op value fill [text])} fills the cells in {it:range} whose
value meets the condition. {it:op} is one of {cmd:==} {cmd:!=} {cmd:>} {cmd:>=}
{cmd:<} {cmd:<=}. {it:value} is a number or text; put text containing spaces in
double quotes, e.g. {cmd:highlight(C2:C164 == "High income" lightblue)}. {it:fill}
is the fill colour and the optional {it:text} is the text colour. Empty cells are
never highlighted, so missing values are left alone even by conditions such as
{cmd:<= 1} (Excel's own comparison rules would treat them as 0).

{phang}
{opt dropdown(range choices)} puts a dropdown list in each cell of {it:range}.
{it:choices} is a comma-separated list, e.g. {cmd:dropdown(F2:F164 Yes,No,Unsure)},
or a reference to cells holding the choices, starting with {cmd:=}, e.g.
{cmd:dropdown(F2:F164 =Lists!A1:A12)}. Do not type {cmd:$} signs: Stata would read
{cmd:$A} as a global macro. {cmd:xlpolish} makes the reference fixed ({cmd:$A$1:$A$12}) itself,
so every dropdown in the range points at the same cells. Typed lists are limited by Excel to
255 characters; use a cell reference for longer lists.

{phang}
{opt iconset(range style [reverse percentile iconsonly])} shows an icon in each
cell according to its value. Styles: {cmd:traffic}, {cmd:trafficrim}, {cmd:arrows},
{cmd:arrowsgray}, {cmd:flags}, {cmd:signs}, {cmd:symbols}, {cmd:symbolsplain}
(three icons); {cmd:arrows4}, {cmd:arrowsgray4}, {cmd:traffic4}, {cmd:rating4},
{cmd:redtoblack4} (four); {cmd:arrows5}, {cmd:arrowsgray5}, {cmd:rating5},
{cmd:quarters5} (five). The range is split into equal shares of the span from
lowest to highest value, as in Excel; add {cmd:percentile} to split by
percentiles instead (equal numbers of cells per icon), {cmd:reverse} to flip
the icon order, and {cmd:iconsonly} to hide the numbers.
E.g. {cmd:iconset(B2:B75 traffic)} or {cmd:iconset(C2:C75 arrows5 percentile)}.

{phang}
{opt top(range N fill [text])} and {opt bottom(range N fill [text])} fill the
{it:N} highest or lowest values, or the top or bottom {it:N} percent with
{cmd:N%}, e.g. {cmd:top(B2:B75 10 lightgreen)} or {cmd:bottom(B2:B75 20% pink)}.

{phang}
{opt average(range above|below fill [text])} fills the values above or below the
average of the range, e.g. {cmd:average(B2:B75 below FFC7CE 9C0006)}.

{pstd}
Like {opt highlight()}, these take several rules separated by semicolons.

{pstd}
{it:Rule order.} When several rules apply to the same cells and both change how
a cell looks (a colour scale and a highlight, say), the rule higher in Excel's
list wins. As when adding rules by hand in Excel, the rules from each
{cmd:xlpolish} call go above the rules already in the sheet, so a highlight
added later shows over an earlier colour scale. Within one call, {opt highlight()},
{opt top()}, {opt bottom()} and {opt average()} go above {opt colorscale()},
{opt databar()} and {opt iconset()}. The order can be changed in Excel under
Home > Conditional Formatting > Manage Rules.

{dlgtab:Data entry}

{phang}
{opt validate(range type rule ["message"])} lets only suitable values be typed
into the cells. {it:type} is {cmd:whole} (whole numbers), {cmd:decimal} (any
number), {cmd:date} (dates written {cmd:YYYY-MM-DD}) or {cmd:length} (number of
characters of text). {it:rule} is {cmd:between} {it:a b}, {cmd:notbetween} {it:a b},
or one of {cmd:==} {cmd:!=} {cmd:>} {cmd:>=} {cmd:<} {cmd:<=} followed by a value.
Empty cells are always allowed. A wrong entry is refused with a message such as
"Enter a number from 0 to 100."; give your own as the last item, in quotes.
Examples: {cmd:validate(B2:B200 decimal between 0 100)},
{cmd:validate(C2:C200 whole >= 1)},
{cmd:validate(D2:D200 date between 2020-01-01 2026-12-31)},
{cmd:validate(F2:F200 length <= 500 "Please keep comments under 500 characters")}.

{phang}
{opt prompt(range "text")} shows a short message whenever one of the cells is
selected, e.g. {cmd:prompt(B2:B200 "Enter the 2025 value in local currency")}.
If the same cells have a dropdown or validation rule, the message is added to it.

{pstd}
Excel allows one dropdown or validation rule per cell. Giving
{opt dropdown()}, {opt validate()} or {opt prompt()} on exactly the same range as
an existing rule replaces it (keeping its message); a range that only partly
overlaps an existing rule is refused, with a message saying which.

{dlgtab:Notes}

{phang}
{opt comment(cell "text")} adds a note to {it:cell}: a small red triangle in the
cell's corner, with the text shown in a pop-up when the mouse is over the cell.
Excel 365 calls these "Notes"; older versions call them comments. Give several
notes separated by semicolons, e.g. {cmd:comment(B1 "Price in US dollars"; C1 "Miles per gallon")};
semicolons inside the quotes are kept as text. Write {cmd:\n} for a line break.
A note on a cell that already has one replaces it, and {cmd:comment(B1 "")}
deletes the note on B1. Notes already in the sheet are kept, with their authors;
any special formatting inside existing note text becomes plain text.

{phang}
{opt commentauthor(name)} sets the author recorded with the new notes. The
default is your user name ({cmd:c(username)}).

{phang}
{opt removecomments} deletes every note on the sheet. It is applied before any
{opt comment()} notes in the same call, so the two together replace all notes.

{dlgtab:Grouping}

{phang}
{opt grouprows(rows)} groups each block of rows, such as {cmd:5:20}, so readers
can fold it away with the {cmd:-} button at the left of the sheet. Groups that
overlap are nested, up to Excel's limit of 7 levels. The button sits on the row
just below each group.

{phang}
{opt groupcols(cols)} does the same for columns, such as {cmd:C:F}; the button
sits above the column just right of each group.

{pmore}
Groups need a gap between them. Excel records grouping only as a level for each
row or column, so two groups that touch, such as {cmd:groupcols(C:E F:H)}, become
one long group, exactly as when grouping by hand in Excel. Leave at least one
ungrouped row or column between groups, e.g. {cmd:groupcols(C:E G:I)}.
{cmd:xlpolish} still applies touching groups but shows a note saying where they
touch. Groups inside other groups (such as {cmd:grouprows(3:12 5:7)}) are fine.

{pmore}
Tip: if the gap column has nothing to hold, make it a slim, empty spacer, e.g.
{cmd:groupcols(C:E G:I) colwidth(F=2)}. A width of 2 keeps the {cmd:-}/{cmd:+}
button fully visible and easy to click; at 1 the column is narrower than the
button. The spacer can be added in Stata before exporting, e.g. with
{cmd:generate byte spacer = .} and {cmd:order spacer, after(}{it:lastvar}{cmd:)}.

{phang}
{opt collapse} starts the groups added in this call folded, so the grouped rows
or columns are hidden until a reader clicks {cmd:+}.

{phang}
{opt ungroup} removes all row and column grouping from the sheet and shows any
rows or columns that a collapsed group had hidden.

{phang}
{opt summaryabove} and {opt summaryleft} change which side of each group its
{cmd:-}/{cmd:+} button sits on, and where Excel expects the group's summary (for
example, a subtotal row or a pillar score column): above a row group instead of
below it, and left of a column group instead of right of it. This suits tables
where a total comes first and its detail follows, e.g. a score column followed
by the indicators behind it: {cmd:groupcols(D:H) summaryleft} puts the button on
column C. {opt summarybelow} and {opt summaryright} switch back to Excel's
default. These are settings for the whole sheet (Excel's "Summary rows below
detail" and "Summary columns to right of detail" boxes), so they apply to all of
its groups. Give them in the same call as {opt collapse}, or before it, so that
collapsed groups are marked on the correct side. Groups still need a gap row or
column between them on either setting.

{dlgtab:Protection and visibility}

{phang}
{opt protect} protects the sheet: readers can select cells, sort, and use filter
dropdowns, but cannot change locked cells, insert or delete rows and columns, or
change formatting. Every cell is locked unless you {opt unlock()} it.

{phang}
{opt password(text)} sets the password needed to remove protection in Excel. Only
a salted SHA-512 fingerprint of the password is stored in the file, the same way
Excel 2013 and later store it; the password itself is not. Note that the password
is visible in your do-file and in Stata's command history. Sheet protection
guards against accidental edits; it is not encryption, and the data in the file
can still be read.

{phang}
{opt unprotect} removes protection, with or without a password.

{phang}
{opt unlock(ranges)} keeps the cells in {it:ranges} editable when the sheet is
protected, as in a data-entry template. Each cell keeps its formatting. Give
rectangles such as {cmd:D2:D500}, whole columns such as {cmd:D:F} (cells typed
into those columns later are editable too), or whole rows such as {cmd:2:5000}.
Rectangles are limited to 2 million cells; use whole columns or rows beyond that.
Unlocking a rectangle stores each of its cells in the file, as Excel does for any
formatted cell, so if the cells are empty, {helpb import excel} will later read
them as part of the table (for example, as an extra empty variable). Unlocking
whole columns or rows stores the setting on the column or row instead and avoids
this.
Dropdown lists made with {opt dropdown()} only work on a protected sheet if
their cells are unlocked.

{phang}
{opt hidesheet} hides the sheet's tab; readers can show it again in Excel by
right-clicking a tab and choosing Unhide. A workbook must keep at least one
visible sheet, so {opt hidesheet} cannot be combined with {opt allsheets}. If
the hidden sheet was the one Excel opens on, the first visible sheet takes its
place. {opt showsheet} shows a hidden sheet again.

{dlgtab:Printing}

{phang}
{opt landscape} and {opt portrait} set the page orientation.

{phang}
{opt fitwidth} scales the printout so all columns fit on one page width; it may
run to as many pages tall as needed.

{phang}
{opt printtitles(rows [cols])} repeats rows (such as {cmd:1} or {cmd:1:2}) and/or
columns (such as {cmd:A} or {cmd:A:B}) on every printed page.

{phang}
{opt header("text")} and {opt footer("text")} set the text printed at the top or
bottom of every page. One section is centred; three sections separated by
{cmd:|} go left, centre and right. These placeholders are filled in by Excel:
{cmd:[page]}, {cmd:[pages]}, {cmd:[date]}, {cmd:[time]}, {cmd:[file]} and
{cmd:[sheet]}. E.g. {cmd:footer("[file] | Page [page] of [pages] | [date]")}.
{cmd:header("")} or {cmd:footer("")} removes it. ({cmd:putexcel} can set headers
and footers from Stata 18; {cmd:xlpolish} also works on Stata 16 and 17.)


{title:Examples}

{pstd}Write a table, then freeze the header, add filters and size the columns{p_end}
{phang2}{cmd:. sysuse auto, clear}{p_end}
{phang2}{cmd:. export excel using "cars.xlsx", firstrow(variables) sheet("Data") replace}{p_end}
{phang2}{cmd:. xlpolish using "cars.xlsx", freeze(B2) autofilter autofit}{p_end}

{pstd}Shade a score column, flag top-quintile cells, and set up printing{p_end}
{phang2}{cmd:. xlpolish using "cars.xlsx", sheet("Data") colorscale(B2:B75) ///}{p_end}
{phang2}{cmd:      highlight(C2:C75 >= 30 lightgreen darkgreen) landscape fitwidth printtitles(1)}{p_end}

{pstd}Same look on every sheet{p_end}
{phang2}{cmd:. xlpolish using "results.xlsx", allsheets freeze(A2) hidegrid zoom(90)}{p_end}

{pstd}A data-entry template with dropdowns{p_end}
{phang2}{cmd:. xlpolish using "template.xlsx", dropdown(D2:D500 Yes,No,Not applicable) tabcolor(orange)}{p_end}

{pstd}Explain columns with notes, and let readers fold away detail rows{p_end}
{phang2}{cmd:. xlpolish using "cars.xlsx", comment(B1 "Price in 1978 US dollars"; C1 "Miles per gallon") ///}{p_end}
{phang2}{cmd:      grouprows(54:75)}{p_end}

{pstd}A protected data-entry template: only the answer cells can be edited, and the list of choices lives on a hidden sheet{p_end}
{phang2}{cmd:. xlpolish using "template.xlsx", sheet("Form") dropdown(B2:B200 =Lists!A1:A3) ///}{p_end}
{phang2}{cmd:      unlock(B2:C200) protect password(team2026)}{p_end}
{phang2}{cmd:. xlpolish using "template.xlsx", sheet("Lists") hidesheet}{p_end}

{pstd}A scored table: icons, top 10, a footer with page numbers, and a hidden helper column{p_end}
{phang2}{cmd:. xlpolish using "scores.xlsx", iconset(C2:C164 traffic) top(B2:B164 10 lightgreen) ///}{p_end}
{phang2}{cmd:      footer("[file] | Page [page] of [pages] | [date]") hidecols(H)}{p_end}

{pstd}Score entry with a range check and an instruction{p_end}
{phang2}{cmd:. xlpolish using "template.xlsx", validate(C2:C200 decimal between 0 100) ///}{p_end}
{phang2}{cmd:      prompt(C2:C200 "Enter a score from 0 to 100") unlock(C2:C200) protect}{p_end}

{pstd}Undo view settings{p_end}
{phang2}{cmd:. xlpolish using "cars.xlsx", freeze(A1) showgrid tabcolor(none) removefilter}{p_end}


{title:Stored results}

{pstd}{cmd:xlpolish} stores the following in {cmd:r()}:

{synoptset 12 tabbed}{...}
{synopt:{cmd:r(file)}}full path of the workbook{p_end}


{title:Remarks}

{pstd}
Large sheets: {opt autofit} and {opt autofilter} read every cell. A sheet of
100,000 rows by 12 columns takes roughly 10-15 seconds. If Java runs out of
memory on a very large sheet, raise Stata's Java memory limit with
{cmd:set java_heapmax}, then restart Stata.

{pstd}
After replacing {cmd:xlpolish.jar} with a newer version, restart Stata, because
Java classes stay loaded for the whole session.


{title:Install and update}

{phang2}{cmd:. net install xlpolish, from("https://raw.githubusercontent.com/fahad-mirza/xlpolish/main/") replace}{p_end}

{pstd}
Run the same line again to update, then restart Stata. Source code, the change
log and issue reports: {browse "https://github.com/fahad-mirza/xlpolish"}.


{title:Requirements}

{pstd}
Stata 16 or newer, on Windows, macOS or Linux. Nothing else: {cmd:xlpolish} is
written in Java and runs on the Java that comes with Stata, so there is no Java to
install, and it does not use Excel, Python, R, PowerShell or macros.
{cmd:xlpolish.jar} must be on the {helpb adopath} next to {cmd:xlpolish.ado};
{cmd:net install} and copying the files into your PERSONAL folder both do this.

{pstd}
A test do-file comes with the package. After {cmd:net install}, fetch it into the
current folder with {cmd:net get xlpolish} (using the same {cmd:from()} as for the
install) and run {cmd:do xlpolish_test.do}; it ends with "ALL TESTS PASSED" and
writes a demo workbook with every feature on to {cmd:xlpolish_test_output}.


{title:Author}

{pstd}
Fahad Mirza{break}
GitHub: {browse "https://github.com/fahad-mirza":github.com/fahad-mirza}

{pstd}
Questions, bug reports and suggestions are welcome at
{browse "https://github.com/fahad-mirza/xlpolish/issues":github.com/fahad-mirza/xlpolish/issues}.
When reporting a problem, please include the {cmd:xlpolish} command you ran and,
if possible, the workbook before it was polished.


{title:Also see}

{psee}
Help: {helpb putexcel}, {helpb export excel}, {helpb import excel}
{p_end}

