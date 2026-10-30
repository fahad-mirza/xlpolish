*! xlpolish 1.2.0  30oct2026
*! Finishing touches for Excel files written by putexcel / export excel:
*! freeze panes, filters, column widths, conditional formatting, dropdowns,
*! notes, grouping, protection, hidden sheets, print setup and sheet appearance.
*! Uses Stata's built-in Java (Stata 16+).
*! Author: Fahad Mirza  https://github.com/fahad-mirza/xlpolish
program define xlpolish, rclass
    version 16

    syntax using/ , [                                                       ///
        sheet(string) allsheets                                             ///
        freeze(string) hidegrid showgrid zoom(integer 0) tabcolor(string)   ///
        autofilter filterrange(string) removefilter                         ///
        autofit autofitmax(real -1) colwidth(string)                        ///
        landscape portrait fitwidth printtitles(string)                     ///
        colorscale(string asis) databar(string asis)                        ///
        highlight(string asis) dropdown(string asis)                        ///
        comment(string asis) commentauthor(string) removecomments           ///
        grouprows(string) groupcols(string) collapse ungroup                ///
        summaryabove summarybelow summaryleft summaryright                  ///
        protect password(string) unprotect unlock(string)                   ///
        hidesheet showsheet                                                 ///
        ]

    * ---- conflicting options ------------------------------------------------
    if `"`sheet'"' != "" & "`allsheets'" != "" {
        display as error "options sheet() and allsheets may not be combined"
        exit 198
    }
    if "`hidegrid'" != "" & "`showgrid'" != "" {
        display as error "options hidegrid and showgrid may not be combined"
        exit 198
    }
    if "`landscape'" != "" & "`portrait'" != "" {
        display as error "options landscape and portrait may not be combined"
        exit 198
    }
    if "`removefilter'" != "" & ("`autofilter'" != "" | `"`filterrange'"' != "") {
        display as error "option removefilter may not be combined with autofilter or filterrange()"
        exit 198
    }
    if `autofitmax' != -1 & "`autofit'" == "" {
        display as error "option autofitmax() requires autofit"
        exit 198
    }
    if "`protect'" != "" & "`unprotect'" != "" {
        display as error "options protect and unprotect may not be combined"
        exit 198
    }
    if `"`password'"' != "" & "`protect'" == "" {
        display as error "option password() requires protect"
        exit 198
    }
    if "`hidesheet'" != "" & "`showsheet'" != "" {
        display as error "options hidesheet and showsheet may not be combined"
        exit 198
    }
    if "`hidesheet'" != "" & "`allsheets'" != "" {
        display as error "options hidesheet and allsheets may not be combined (a workbook needs a visible sheet)"
        exit 198
    }
    if "`collapse'" != "" & `"`grouprows'`groupcols'"' == "" {
        display as error "option collapse requires grouprows() or groupcols()"
        exit 198
    }
    if "`ungroup'" != "" & `"`grouprows'`groupcols'"' != "" {
        display as error "option ungroup may not be combined with grouprows() or groupcols()"
        exit 198
    }
    if "`summaryabove'" != "" & "`summarybelow'" != "" {
        display as error "options summaryabove and summarybelow may not be combined"
        exit 198
    }
    if "`summaryleft'" != "" & "`summaryright'" != "" {
        display as error "options summaryleft and summaryright may not be combined"
        exit 198
    }
    if `"`commentauthor'"' == "" local commentauthor `"`c(username)'"'
    if `zoom' != 0 & !inrange(`zoom', 10, 400) {
        display as error "zoom() must be between 10 and 400"
        exit 198
    }

    * ---- file: Java does not follow Stata's -cd-, so pass an absolute path --
    local file `"`using'"'
    local c1 = substr(`"`file'"', 1, 1)
    local c2 = substr(`"`file'"', 2, 1)
    if !(`"`c1'"' == "/" | `"`c1'"' == "\" | `"`c2'"' == ":") {
        local file `"`c(pwd)'/`file'"'
    }
    confirm file `"`file'"'
    if !ustrregexm(strlower(`"`file'"'), "\.xls[xm]$") {
        display as error "xlpolish works with .xlsx and .xlsm files only (not old .xls files)"
        exit 198
    }

    * ---- write the options to a small text file for the Java side ----------
    *      (one "key<TAB>value" per line, so values may hold spaces and quotes)
    tempfile spec
    tempname fh
    file open `fh' using `"`spec'"', write text
    file write `fh' "file" _tab `"`file'"' _n
    if `"`sheet'"'       != "" file write `fh' "sheet"       _tab `"`sheet'"' _n
    if "`allsheets'"     != "" file write `fh' "allsheets"   _tab "1" _n
    if `"`freeze'"'      != "" file write `fh' "freeze"      _tab `"`freeze'"' _n
    if "`hidegrid'"      != "" file write `fh' "gridlines"   _tab "off" _n
    if "`showgrid'"      != "" file write `fh' "gridlines"   _tab "on" _n
    if `zoom'            != 0  file write `fh' "zoom"        _tab "`zoom'" _n
    if `"`tabcolor'"'    != "" file write `fh' "tabcolor"    _tab `"`tabcolor'"' _n
    if "`autofilter'"    != "" file write `fh' "autofilter"  _tab "on" _n
    if `"`filterrange'"' != "" file write `fh' "filterrange" _tab `"`filterrange'"' _n
    if "`removefilter'"  != "" file write `fh' "autofilter"  _tab "off" _n
    if "`autofit'"       != "" file write `fh' "autofit"     _tab "1" _n
    if `autofitmax'      != -1 file write `fh' "autofitmax"  _tab "`autofitmax'" _n
    if `"`colwidth'"'    != "" file write `fh' "colwidth"    _tab `"`colwidth'"' _n
    if "`landscape'"     != "" file write `fh' "orientation" _tab "landscape" _n
    if "`portrait'"      != "" file write `fh' "orientation" _tab "portrait" _n
    if "`fitwidth'"      != "" file write `fh' "fitwidth"    _tab "1" _n
    if `"`printtitles'"' != "" file write `fh' "printtitles" _tab `"`printtitles'"' _n
    if `"`colorscale'"'  != "" file write `fh' "colorscale"  _tab `"`colorscale'"' _n
    if `"`databar'"'     != "" file write `fh' "databar"     _tab `"`databar'"' _n
    if `"`highlight'"'   != "" file write `fh' "highlight"   _tab `"`highlight'"' _n
    if `"`dropdown'"'    != "" file write `fh' "dropdown"    _tab `"`dropdown'"' _n
    if `"`comment'"'     != "" file write `fh' "comment"     _tab `"`comment'"' _n
    if `"`comment'"'     != "" file write `fh' "commentauthor" _tab `"`commentauthor'"' _n
    if "`removecomments'" != "" file write `fh' "removecomments" _tab "1" _n
    if `"`grouprows'"'   != "" file write `fh' "grouprows"   _tab `"`grouprows'"' _n
    if `"`groupcols'"'   != "" file write `fh' "groupcols"   _tab `"`groupcols'"' _n
    if "`collapse'"      != "" file write `fh' "collapse"    _tab "1" _n
    if "`ungroup'"       != "" file write `fh' "ungroup"     _tab "1" _n
    if "`summaryabove'"  != "" file write `fh' "summaryrows" _tab "above" _n
    if "`summarybelow'"  != "" file write `fh' "summaryrows" _tab "below" _n
    if "`summaryleft'"   != "" file write `fh' "summarycols" _tab "left" _n
    if "`summaryright'"  != "" file write `fh' "summarycols" _tab "right" _n
    if `"`unlock'"'      != "" file write `fh' "unlock"      _tab `"`unlock'"' _n
    if "`protect'"       != "" file write `fh' "protect"     _tab "1" _n
    if `"`password'"'    != "" file write `fh' "password"    _tab `"`password'"' _n
    if "`unprotect'"     != "" file write `fh' "unprotect"   _tab "1" _n
    if "`hidesheet'"     != "" file write `fh' "hidesheet"   _tab "1" _n
    if "`showsheet'"     != "" file write `fh' "showsheet"   _tab "1" _n
    file close `fh'

    capture noisily javacall xlpolish.XLPolish run, jars(xlpolish.jar) args(`"`spec'"')
    local rc = _rc
    * the option file may hold a password: remove it straight away
    capture erase `"`spec'"'
    if `rc' exit `rc'

    return local file `"`file'"'
end
