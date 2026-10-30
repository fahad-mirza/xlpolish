* xlpolish_test.do -- end-to-end checks for xlpolish in Stata
*
* Two ways to run it:
*   1. From a downloaded copy of the package: set XLP_KIT below to that folder.
*      Excel outputs go to a test_output subfolder of it.
*   2. After -net install xlpolish- (get this file with -net get xlpolish-):
*      leave XLP_KIT empty. Excel outputs go to xlpolish_test_output in the
*      current folder.
* It should end with "ALL TESTS PASSED".
* Restart Stata first if an older xlpolish.jar was used in this session.

version 16
clear all
set more off

global XLP_KIT ""          // e.g. "C:/Users/you/Documents/xlpolish"; leave empty after net install

if `"$XLP_KIT"' != "" {
    adopath ++ "$XLP_KIT"
    local workdir `"$XLP_KIT/test_output"'
}
else {
    capture which xlpolish
    if _rc {
        display as error "xlpolish is not installed; install it with net install, or set XLP_KIT to the package folder"
        exit 111
    }
    local workdir `"`c(pwd)'/xlpolish_test_output"'
}
which xlpolish
capture mkdir `"`workdir'"'
local f `"`workdir'/xlpolish_test.xlsx"'

* scratch folder for reading the XML inside the .xlsx (no trailing slash)
global XLP_TMP `"`c(tmpdir)'"'
if inlist(substr(`"$XLP_TMP"', -1, 1), "/", "\") global XLP_TMP = substr(`"$XLP_TMP"', 1, strlen(`"$XLP_TMP"') - 1)

* ---------------------------------------------------------------------------
* Helpers: load a part of an .xlsx (e.g. xl/worksheets/sheet1.xml) into the
* string scalar __xlp_xml, and check text is (or is not) in it.
* The Nth sheet created below is stored as sheetN.xml.
* ---------------------------------------------------------------------------
capture program drop _xlp_xml
program define _xlp_xml
    args file part
    local here `"`c(pwd)'"'
    local tmp  `"$XLP_TMP/xlpolish_unzip"'
    capture mkdir `"`tmp'"'
    quietly cd `"`tmp'"'
    capture erase "`part'"
    quietly copy `"`file'"' "book.zip", replace
    quietly unzipfile "book.zip", replace
    scalar __xlp_xml = fileread("`part'")
    quietly cd `"`here'"'
end
capture program drop _has
program define _has
    * _has "text" [0]  -> assert text is (or with 0, is not) in __xlp_xml
    args text want
    if "`want'" == "" local want 1
    if `want' assert strpos(__xlp_xml, `"`text'"') > 0
    else      assert strpos(__xlp_xml, `"`text'"') == 0
end

* ---------------------------------------------------------------------------
* Build a 3-sheet workbook with export excel + putexcel
* ---------------------------------------------------------------------------
sysuse auto, clear
export excel using `"`f'"', sheet("Data") firstrow(variables) nolabel replace

putexcel set `"`f'"', sheet("Region's summary") modify
putexcel A1 = "Region" B1 = "Count", bold
putexcel A2 = "East Asia & Pacific" B2 = 25
putexcel A3 = "South Asia" B3 = 8

putexcel set `"`f'"', sheet("Formatted") modify
putexcel A1:D1 = "Merged header", merge hcenter bold
putexcel A2 = "x" B2 = 1.5, nformat(number_d2)
putexcel clear

import excel using `"`f'"', sheet("Data") firstrow clear
tempfile before
save `before'

* ---------------------------------------------------------------------------
* 1. View, data, conditional formatting, dropdown and print options together
*    (Data is A1:L75 = header + 74 cars)
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("Data") freeze(B2) hidegrid zoom(90) tabcolor(navy) ///
    autofilter autofit colwidth(A=30) landscape fitwidth printtitles(1)          ///
    colorscale(B2:B75) databar(C2:C75 steelblue)                                ///
    highlight(D2:D75 >= 4 lightgreen darkgreen; D2:D75 <= 1 FFC7CE 9C0006)       ///
    dropdown(M2:M75 Yes, No, Unsure)
assert r(file) == `"`f'"'
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has `"state="frozen""'
_has `"topLeftCell="B2""'
_has `"showGridLines="0""'
_has `"zoomScale="90""'
_has `"rgb="FF000080""'
_has `"<autoFilter ref="A1:L75""'
_has `"width="30""'
_has `"orientation="landscape""'
_has `"fitToPage="1""'
_has `"type="colorScale""'
_has `"type="dataBar""'
_has "NOT(ISBLANK(D2)),D2&gt;=4"
_has "NOT(ISBLANK(D2)),D2&lt;=1"
_has `"sqref="M2:M75""'
_has `""Yes,No,Unsure""'
_xlp_xml `"`f'"' "xl/workbook.xml"
_has "_xlnm.Print_Titles"
_has "_xlnm._FilterDatabase"
display as result "PASS 1: view, data, formatting, dropdown and print options"

* ---------------------------------------------------------------------------
* 2. Sheet name with an apostrophe; text value with spaces and &
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("region's SUMMARY") freeze(A2) ///
    highlight(A2:A3 == "East Asia & Pacific" lightblue) printtitles(1)
_xlp_xml `"`f'"' "xl/worksheets/sheet2.xml"
_has `"topLeftCell="A2""'
_has "East Asia &amp; Pacific"
_xlp_xml `"`f'"' "xl/workbook.xml"
_has "'Region''s summary'!"
display as result "PASS 2: apostrophe sheet name, quoted text value"

* ---------------------------------------------------------------------------
* 3. allsheets
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', allsheets zoom(120)
forvalues s = 1/3 {
    _xlp_xml `"`f'"' "xl/worksheets/sheet`s'.xml"
    _has `"zoomScale="120""'
}
display as result "PASS 3: allsheets"

* ---------------------------------------------------------------------------
* 4. Undo view settings on one sheet only
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("Data") freeze(A1) showgrid tabcolor(none) removefilter
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has `"state="frozen""' 0
_has "showGridLines" 0
_has "tabColor" 0
_has "<autoFilter" 0
_has `"type="colorScale""'
_xlp_xml `"`f'"' "xl/worksheets/sheet2.xml"
_has `"state="frozen""'
display as result "PASS 4: removal options"

* ---------------------------------------------------------------------------
* 5. Relative path after cd
* ---------------------------------------------------------------------------
local here `"`c(pwd)'"'
quietly cd `"`workdir'"'
xlpolish using "xlpolish_test.xlsx", sheet("Formatted") tabcolor("31 78 121") autofit
quietly cd `"`here'"'
_xlp_xml `"`f'"' "xl/worksheets/sheet3.xml"
_has `"rgb="FF1F4E79""'
_has `"customWidth="1""'
display as result "PASS 5: relative path, RGB colour, autofit with merged header"

* ---------------------------------------------------------------------------
* 6. Notes, grouping, unlocked cells, protection with a password
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("Data")                                          ///
    comment(B1 "Price in 1978 US dollars"; C1 "Miles per gallon; city") commentauthor("Test") ///
    grouprows(10:20) groupcols(E:F) collapse unlock(M2:M75) protect password(test123)
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has "<legacyDrawing"
_has `"outlineLevel="1""'
_has `"collapsed="1""'
_has `"algorithmName="SHA-512""'
_has "test123" 0
_xlp_xml `"`f'"' "xl/comments1.xml"
_has "Miles per gallon; city"
_has "<author>Test</author>"
_xlp_xml `"`f'"' "xl/styles.xml"
_has `"locked="0""'
display as result "PASS 6: notes, grouping, unlocked cells, password protection"

* ---------------------------------------------------------------------------
* 7. Undo them; hide and show a sheet
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("Data") removecomments ungroup unprotect
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has "<legacyDrawing" 0
_has "outlineLevel" 0
_has "sheetProtection" 0
xlpolish using `"`f'"', sheet("Formatted") hidesheet
_xlp_xml `"`f'"' "xl/workbook.xml"
_has `"state="hidden""'
xlpolish using `"`f'"', sheet("Formatted") showsheet
_xlp_xml `"`f'"' "xl/workbook.xml"
_has `"state="hidden""' 0
xlpolish using `"`f'"', sheet("Region's summary") grouprows(2:3) collapse summaryabove
_xlp_xml `"`f'"' "xl/worksheets/sheet2.xml"
_has `"summaryBelow="0""'
_has `"collapsed="1""'
xlpolish using `"`f'"', sheet("Region's summary") ungroup summarybelow
_xlp_xml `"`f'"' "xl/worksheets/sheet2.xml"
_has "summaryBelow" 0
display as result "PASS 7: removal of notes, grouping, protection; hide and show; summary side"

* ---------------------------------------------------------------------------
* 7b. Filter arrows, validation and prompts, icons and rank rules,
*     header/footer, hidden columns
* ---------------------------------------------------------------------------
xlpolish using `"`f'"', sheet("Data") autofilter nofilterarrow(L)               ///
    validate(N2:N75 decimal between 0 100) prompt(N2:N75 "Score from 0 to 100") ///
    iconset(B2:B75 traffic) top(C2:C75 5 lightgreen) average(B2:B75 below pink) ///
    header("xlpolish test") footer("Page [page] of [pages]") hidecols(K)
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has `"hiddenButton="1""'
_has `"type="decimal""'
_has `"prompt="Score from 0 to 100""'
_has `"iconSet="3TrafficLights1""'
_has `"type="top10""'
_has `"priority="1" rank="5" type="top10""'
_has `"aboveAverage="0""'
_has "Page &amp;P of &amp;N"
_has `"hidden="1""'
xlpolish using `"`f'"', sheet("Data") removefilter showcols(K) header("") footer("")
_xlp_xml `"`f'"' "xl/worksheets/sheet1.xml"
_has "hiddenButton" 0
_has "headerFooter" 0
_has `"hidden="1""' 0
display as result "PASS 7b: filter arrows, validation, prompts, icons, rank rules, header/footer, hidden columns"

* ---------------------------------------------------------------------------
* 8. Data unchanged; other sheets still read
* ---------------------------------------------------------------------------
import excel using `"`f'"', sheet("Data") firstrow clear
* unlock(M2:M75) in test 6 stored those empty cells (with their unlocked
* formatting) in the file, so import excel reads an extra, completely empty
* column M. Check it is empty, then compare the original variables.
capture confirm variable M
if !_rc {
    assert missing(M)
    drop M
}
cf _all using `before'
import excel using `"`f'"', sheet("Formatted") cellrange(A2:B2) clear
assert B == 1.5
display as result "PASS 8: data intact"

* ---------------------------------------------------------------------------
* 9. Errors exit with the right codes and leave the file untouched
* ---------------------------------------------------------------------------
checksum `"`f'"'
local ck = r(checksum)
rcof `"noisily xlpolish using "`f'", sheet("Data") allsheets zoom(90)"'         == 198
rcof `"noisily xlpolish using "`f'", hidegrid showgrid"'                        == 198
rcof `"noisily xlpolish using "`f'", landscape portrait"'                       == 198
rcof `"noisily xlpolish using "`f'", autofilter removefilter"'                  == 198
rcof `"noisily xlpolish using "`f'", autofitmax(40)"'                           == 198
rcof `"noisily xlpolish using "`f'", zoom(5)"'                                  == 198
rcof `"noisily xlpolish using "`f'""'                                           == 198
rcof `"noisily xlpolish using "`f'", tabcolor(notacolour)"'                     == 198
rcof `"noisily xlpolish using "`f'", zoom(90) highlight(B2:B9 === 5 red)"'      == 198
rcof `"noisily xlpolish using "`f'", sheet("Nope") zoom(90)"'                   == 198
rcof `"noisily xlpolish using "`f'", protect unprotect"'                        == 198
rcof `"noisily xlpolish using "`f'", password(x) zoom(90)"'                     == 198
rcof `"noisily xlpolish using "`f'", hidesheet showsheet"'                      == 198
rcof `"noisily xlpolish using "`f'", allsheets hidesheet"'                      == 198
rcof `"noisily xlpolish using "`f'", collapse zoom(90)"'                        == 198
rcof `"noisily xlpolish using "`f'", ungroup grouprows(2:3)"'                   == 198
rcof `"noisily xlpolish using "`f'", summaryabove summarybelow"'                == 198
rcof `"noisily xlpolish using "`f'", summaryleft summaryright"'                 == 198
rcof `"noisily xlpolish using "`f'", autofilter removefilter nofilterarrow(E)"' == 198
rcof `"noisily xlpolish using "`f'", validate(B2:B9 decimal between 9 1)"'      == 198
rcof `"noisily xlpolish using "`f'", footer("Page [pg]")"'                      == 198
rcof `"noisily xlpolish using "`f'", zoom(90) comment(B1:C2 "range not cell")"' == 198
rcof `"noisily xlpolish using "`workdir'/no_such_file.xlsx", zoom(90)"'         == 601
quietly copy `"`f'"' `"`workdir'/old_format.xls"', replace
rcof `"noisily xlpolish using "`workdir'/old_format.xls", zoom(90)"'            == 198
checksum `"`f'"'
assert r(checksum) == `ck'
capture erase `"`workdir'/old_format.xls"'
display as result "PASS 9: errors, file untouched"

* ---------------------------------------------------------------------------
* 10. Demo workbook with every feature left on, for checking by eye in Excel
*     Data     - view, filter, widths, formatting, dropdown, notes, grouping, print
*     Template - protected form (password: demo); only B2:D6 editable,
*                dropdowns take their choices from the hidden Lists sheet,
*                scores in D2:D6 must be 0-100, prompts on B and D
*     Lists    - hidden
*     Scores   - two column groups side by side with a slim spacer column E,
*                a filter on the whole table but no filter arrow on E
* ---------------------------------------------------------------------------
local demo `"`workdir'/xlpolish_demo.xlsx"'
sysuse auto, clear
export excel using `"`demo'"', firstrow(variables) sheet("Data") replace
putexcel set `"`demo'"', sheet("Template") modify
putexcel A1 = "Economy" B1 = "Reviewed?" C1 = "Reviewer comment" D1 = "Score (0-100)", bold
putexcel A2 = "Albania" A3 = "Bangladesh" A4 = "Colombia" A5 = "Ghana" A6 = "Viet Nam"
putexcel set `"`demo'"', sheet("Lists") modify
putexcel A1 = "Yes" A2 = "No" A3 = "Unsure"
* Scores: two blocks of (made-up) indicator columns with an empty spacer column E
putexcel set `"`demo'"', sheet("Scores") modify
putexcel A1 = "Economy" B1 = "Ind 1.1" C1 = "Ind 1.2" D1 = "Ind 1.3" ///
         F1 = "Ind 2.1" G1 = "Ind 2.2" H1 = "Ind 2.3", bold
putexcel A2 = "Economy A" A3 = "Economy B" A4 = "Economy C" ///
         A5 = "Economy D" A6 = "Economy E" A7 = "Economy F"
matrix P1 = (62, 71, 55 \ 80, 77, 69 \ 45, 52, 61 \ 73, 68, 79 \ 57, 49, 63 \ 88, 84, 76)
matrix P2 = (48, 66, 59 \ 72, 81, 64 \ 58, 40, 47 \ 66, 70, 75 \ 61, 54, 68 \ 79, 83, 86)
putexcel B2 = matrix(P1) F2 = matrix(P2)
putexcel clear

xlpolish using `"`demo'"', sheet("Data") freeze(B2) autofilter autofit hidegrid zoom(90) ///
    tabcolor(navy) colorscale(B2:B75) databar(G2:G75)                                 ///
    highlight(C2:C75 >= 30 lightgreen darkgreen; D2:D75 <= 2 FFC7CE 9C0006)            ///
    dropdown(M2:M75 Yes,No) landscape fitwidth printtitles(1)                         ///
    comment(B1 "Price in 1978 US dollars"; C1 "Miles per gallon") grouprows(54:75) ///
    footer("[file] | Page [page] of [pages] | [date]")
xlpolish using `"`demo'"', sheet("Template") freeze(A2) autofit colwidth(C=40) tabcolor(orange) ///
    dropdown(B2:B6 =Lists!A1:A3) validate(D2:D6 decimal between 0 100)                   ///
    prompt(B2:B6 "Choose Yes, No or Unsure"; D2:D6 "Enter a score from 0 to 100")        ///
    unlock(B2:D6) protect password(demo)
xlpolish using `"`demo'"', sheet("Scores") freeze(B2) autofilter autofit nofilterarrow(E) ///
    groupcols(B:D F:H) colwidth(E=2) tabcolor(teal)
xlpolish using `"`demo'"', sheet("Lists") hidesheet
_xlp_xml `"`demo'"' "xl/worksheets/sheet1.xml"
_has `"state="frozen""'
_has `"<autoFilter ref="A1:L75""'
_has "<legacyDrawing"
_xlp_xml `"`demo'"' "xl/worksheets/sheet2.xml"
* (char(36) is the dollar sign; typing it here would start a Stata global macro)
assert strpos(__xlp_xml, "Lists!" + char(36) + "A" + char(36) + "1:" + char(36) + "A" + char(36) + "3") > 0
_has "sheetProtection"
_has "Enter a score from 0 to 100"
_xlp_xml `"`demo'"' "xl/worksheets/sheet4.xml"
_has `"hiddenButton="1""'
_has `"outlineLevel="1""'
_xlp_xml `"`demo'"' "xl/workbook.xml"
_has `"state="hidden""'
display as result "PASS 10: demo workbook written"

scalar drop __xlp_xml
display _newline as result "ALL TESTS PASSED"
display as text "Excel files are in: " as result `"`workdir'"'
display as text "  xlpolish_demo.xlsx - every feature on; check it by eye in Excel (list in README.md)"
display as text "  xlpolish_test.xlsx - left as the tests finish with it (some settings removed on purpose)"
