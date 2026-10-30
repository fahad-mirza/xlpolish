"""Core options: view, data, formatting, dropdowns, printing, errors. Run via run_tests.sh."""
import subprocess, zipfile, shutil, os, re, hashlib, tempfile, warnings
warnings.filterwarnings("ignore")
import openpyxl, xlsxwriter
from openpyxl.chart import BarChart, Reference
from xml.etree import ElementTree as ET

CP = os.environ["XLPOLISH_CP"]  # set by run_tests.sh: the jar, plus the compiled SFI stub and Run.class
ORDER = ["sheetPr","dimension","sheetViews","sheetFormatPr","cols","sheetData","sheetCalcPr","sheetProtection",
 "protectedRanges","scenarios","autoFilter","sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr",
 "conditionalFormatting","dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter",
 "rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags","drawing","legacyDrawing",
 "legacyDrawingHF","drawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"]

def polish(path, **opts):
    """opts mimic what the ado writes: key -> value (lists allowed)"""
    lines = ["file\t" + os.path.abspath(path)]
    for k, v in opts.items():
        for item in (v if isinstance(v, list) else [v]):
            lines.append(f"{k}\t{item}")
    fd, spec = tempfile.mkstemp(suffix=".txt"); os.close(fd)
    open(spec, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    out = subprocess.run(["java","-cp",CP,"Run",spec],capture_output=True,text=True).stdout
    rc = int(re.search(r"RC=(\d+)", out).group(1))
    return rc, out.replace("RC=%d" % rc, "").strip()

fails = 0
def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else "   >> " + str(extra)[:600]))
    if not cond: fails += 1

def order_ok(path):
    """Every worksheet's children appear in schema order."""
    bad = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                root = ET.fromstring(z.read(n))
                ranks = [ORDER.index(c.tag.split('}')[1]) for c in root if c.tag.split('}')[1] in ORDER]
                if ranks != sorted(ranks): bad.append(n)
    return bad

def md5(p): return hashlib.md5(open(p,"rb").read()).hexdigest()

# ---------------------------------------------------------------- workbook 1
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Data"
ws.append(["Economy", "Score", "Rank", "Quintile", "Reviewed", "Long note column"])
for i in range(1, 51):
    ws.append([f"Country {i}" if i != 7 else "Côte d'Ivoire", 50 + i * 0.731, i, (i % 5) + 1, None,
               "short" if i != 3 else "a considerably longer piece of text in this cell"])
ws.merge_cells("H1:J1"); ws["H1"] = "A very long merged title that should not drive widths"
ws2 = wb.create_sheet("Region's summary"); ws2.append(["Region", "N"]); ws2.append(["EAP", 25])
cs = wb.create_chartsheet("Chart"); ch = BarChart(); ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=10)); cs.add_chart(ch)
wb.save("w1.xlsx"); f1 = os.path.abspath("w1.xlsx")

rc, out = polish(f1, sheet="Data", freeze="B2", gridlines="off", zoom="90", tabcolor="31 78 121",
                 autofilter="on", autofit="1", colwidth="A=30", orientation="landscape", fitwidth="1",
                 printtitles="1", colorscale="B2:B51", databar="C2:C51 steelblue",
                 highlight="D2:D51 == 5 lightgreen darkgreen; D2:D51 < 2 FFC7CE 9C0006",
                 dropdown="E2:E51 Yes, No, Maybe")
print("  ", out)
check("all options rc=0", rc == 0, out)
wb = openpyxl.load_workbook(f1); s = wb["Data"]
check("freeze B2", s.freeze_panes == "B2")
check("gridlines hidden", s.sheet_view.showGridLines is False)
check("zoom 90", s.sheet_view.zoomScale == 90)
check("tab colour from R G B", s.sheet_properties.tabColor.rgb == "FF1F4E79", s.sheet_properties.tabColor)
check("autofilter on used range (merged title text is in H1)", s.auto_filter.ref == "A1:H51", s.auto_filter.ref)
check("colwidth A=30 overrides autofit", s.column_dimensions["A"].width == 30, s.column_dimensions["A"].width)
wF = s.column_dimensions["F"].width; wB = s.column_dimensions["B"].width
check("autofit: long text column wide, number column narrow", 40 < wF <= 60 and 6 <= wB < 15, (wF, wB))
check("merged title ignored for widths", s.column_dimensions["H"].width in (None, 13.0) or s.column_dimensions["H"].width < 20,
      s.column_dimensions["H"].width)
check("landscape + fit to 1 page wide", s.page_setup.orientation == "landscape" and s.page_setup.fitToWidth == 1
      and s.page_setup.fitToHeight == 0 and s.sheet_properties.pageSetUpPr.fitToPage)
check("print titles row 1", s.print_title_rows.replace("$","") == "1:1", s.print_title_rows)
rules = [(str(r.sqref), rule.type, rule.operator) for r in s.conditional_formatting for rule in r.rules]
check("4 conditional formats with right types", sorted(t for _, t, _ in rules) == ["colorScale","dataBar","expression","expression"], rules)
prios = [rule.priority for r in s.conditional_formatting for rule in r.rules]
fx = sorted(rule.formula[0] for r in s.conditional_formatting for rule in r.rules if rule.type == "expression")
check("highlight formulas skip blanks", fx == ["AND(NOT(ISBLANK(D2)),D2<2)", "AND(NOT(ISBLANK(D2)),D2=5)"], fx)
check("priorities unique", len(set(prios)) == len(prios), prios)
dxf_fills = [d.fill.bgColor.rgb for d in wb._differential_styles.dxf]
check("highlight fills stored as dxf", "FF90EE90" in dxf_fills and "FFFFC7CE" in dxf_fills, dxf_fills)
dv = s.data_validations.dataValidation
check("dropdown list", len(dv) == 1 and dv[0].formula1 == '"Yes,No,Maybe"' and str(dv[0].sqref) == "E2:E51", [(d.formula1, str(d.sqref)) for d in dv])
check("other sheet untouched", wb["Region's summary"].freeze_panes is None)
check("schema order in all sheets", order_ok(f1) == [], order_ok(f1))
names = zipfile.ZipFile(f1).read("xl/workbook.xml").decode()
check("filter + print-title names in workbook.xml", "_xlnm._FilterDatabase" in names and "_xlnm.Print_Titles" in names
      and "'Data'!$A$1:$H$51" in names and "'Data'!$1:$1" in names, re.findall(r"<definedName.*?</definedName>", names))

# second call accumulates rules with fresh priorities; works on the apostrophe sheet
rc, out = polish(f1, sheet="region's SUMMARY", highlight="B2 > 10 yellow", printtitles="1:1 A", freeze="A2"); print("  ", out)
wb = openpyxl.load_workbook(f1); s2 = wb["Region's summary"]
check("case-insensitive sheet, apostrophe name in print titles", rc == 0 and s2.print_title_rows.replace("$","") == "1:1"
      and s2.print_title_cols.replace("$","") == "A:A", (s2.print_title_rows, s2.print_title_cols))
check("Data sheet still intact after 2nd call", wb["Data"].freeze_panes == "B2" and sum(len(r.rules) for r in wb["Data"].conditional_formatting) == 4)

rc, out = polish(f1, sheet="Data", highlight="C2:C51 >= 40 orange"); 
s = openpyxl.load_workbook(f1)["Data"]; prios = [rule.priority for r in s.conditional_formatting for rule in r.rules]
check("repeat call: 5 rules, unique priorities", len(prios) == 5 and len(set(prios)) == 5, prios)

# removal paths
rc, out = polish(f1, sheet="Data", autofilter="off", freeze="A1", tabcolor="none", gridlines="on"); print("  ", out)
wb = openpyxl.load_workbook(f1); s = wb["Data"]
names = zipfile.ZipFile(f1).read("xl/workbook.xml").decode()
check("filter off + freeze off + tab colour off + gridlines on", rc == 0 and s.auto_filter.ref is None and s.freeze_panes is None
      and s.sheet_properties.tabColor is None and s.sheet_view.showGridLines in (None, True)
      and "_FilterDatabase" not in names and "Print_Titles" in names)

# allsheets skips the chart sheet
rc, out = polish(f1, allsheets="1", zoom="120"); print("  ", out)
wb = openpyxl.load_workbook(f1)
check("allsheets (chart sheet skipped)", rc == 0 and wb["Data"].sheet_view.zoomScale == 120 and wb["Region's summary"].sheet_view.zoomScale == 120)
check("schema order after many edits", order_ok(f1) == [], order_ok(f1))

# ---------------------------------------------------------------- errors leave file untouched
before = md5(f1)
cases = [
    ({"zoom": "500"}, 198), ({"tabcolor": "notacolour"}, 198), ({"colorscale": "B2:ZZZZ9"}, 198),
    ({"highlight": "B2:B5 === 5 red"}, 198), ({"highlight": "B2:B5 == 5"}, 198),
    ({"dropdown": "E2:E5 " + ",".join(["choice%03d" % i for i in range(40)])}, 198),
    ({"sheet": "Nope", "zoom": "90"}, 198), ({"sheet": "Chart", "zoom": "90"}, 198),
    ({"colwidth": "A=300"}, 198), ({"freeze": "2B"}, 198), ({"printtitles": "x1"}, 198),
    ({"filterrange": "A1:J1"}, 198), ({}, 198),
    ({"zoom": "90", "highlight": "B2 == 1 nocolour"}, 198),   # valid zoom + bad rule: nothing applied
]
for opts, want in cases:
    rc, out = polish(f1, **opts)
    check(f"error {opts} -> {want}", rc == want, out)
check("file unchanged after all errors", md5(f1) == before)
rc, out = polish("/tmp/xp/missing.xlsx", zoom="90"); check("missing file -> 601", rc == 601, out)
open("bad.xlsx","w").write("x"); rc, out = polish("bad.xlsx", zoom="90"); check("not a workbook -> 610", rc == 610, out)

# ---------------------------------------------------------------- workbook 2: xlsxwriter, col range splitting, table
w = xlsxwriter.Workbook("w2.xlsx"); a = w.add_worksheet("Table 1"); a.set_column("A:Z", 15)
for r in range(10): a.write_row(r, 0, [f"h{c}" if r == 0 else r * c for c in range(26)])
t = w.add_worksheet("WithTable"); t.add_table("A1:C5", {"data": [[1,2,3]] * 4}); w.close()
f2 = os.path.abspath("w2.xlsx")
rc, out = polish(f2, sheet="Table 1", colwidth="C=40 E:F=8", autofilter="on"); print("  ", out)
s = openpyxl.load_workbook(f2)["Table 1"]
cols = re.findall(r'<col [^>]*/>', zipfile.ZipFile(f2).read("xl/worksheets/sheet1.xml").decode())
ws_ = {c: s.column_dimensions[c].width for c in "ABCDEFGZ"}
spans = [tuple(re.search(p, c).group(1) for p in (r'min="(\d+)"', r'max="(\d+)"', r'width="([\d.]+)"')) for c in cols]
check("column range split correctly", spans == [("1","2","15.7109375"),("3","3","40"),("4","4","15.7109375"),("5","5","8"),("6","6","8"),("7","26","15.7109375")], spans)
check("autofilter on xlsxwriter sheet", s.auto_filter.ref == "A1:Z10", s.auto_filter.ref)
rc, out = polish(f2, sheet="WithTable", autofilter="on"); check("autofilter refused on sheet with Excel table", rc == 198, out)

# ---------------------------------------------------------------- workbook 3: prefixed namespace, no sheetViews, extLst present
with zipfile.ZipFile(f2) as zin, zipfile.ZipFile("w3.xlsx", "w") as zout:
    for n in zin.namelist():
        d = zin.read(n)
        if n == "xl/worksheets/sheet1.xml":
            d = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 b'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                 b'<x:dimension ref="A1:B3"/><x:sheetFormatPr defaultRowHeight="15"/>'
                 b'<x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>V</x:t></x:is></x:c></x:row>'
                 b'<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t xml:space="preserve"> padded </x:t></x:is></x:c><x:c r="B2"><x:v>3</x:v></x:c></x:row>'
                 b'<x:row r="3"><x:c t="inlineStr"><x:is><x:t>no r attr</x:t></x:is></x:c><x:c><x:v>4</x:v></x:c></x:row></x:sheetData>'
                 b'<x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                 b'<x:extLst><x:ext uri="{x}"/></x:extLst></x:worksheet>')
        zout.writestr(n, d)
f3 = os.path.abspath("w3.xlsx")
rc, out = polish(f3, sheet="Table 1", freeze="A2", zoom="80", tabcolor="red", autofit="1", autofilter="on",
                 orientation="portrait", fitwidth="1", colorscale="B2:B3 white red", dropdown="A2:A3 =\'Table 1\'!$A$1:$A$3"); print("  ", out)
x = zipfile.ZipFile(f3).read("xl/worksheets/sheet1.xml").decode()
s = openpyxl.load_workbook(f3)["Table 1"]
check("prefixed sheet: all parts in schema order", rc == 0 and order_ok(f3) == [], (out, order_ok(f3), x[:900]))
check("prefixed sheet: openpyxl reads freeze/filter/cf/dv", s.freeze_panes == "A2" and s.auto_filter.ref == "A1:B3"
      and len(list(s.conditional_formatting)) == 1 and s.data_validations.dataValidation[0].formula1 == "'Table 1'!$A$1:$A$3", x[:900])
check("cells without r attribute handled", s.auto_filter.ref == "A1:B3" and s["A2"].value == " padded ")

# ---------------------------------------------------------------- LibreOffice round trip (independent reader/writer)
os.makedirs("lo", exist_ok=True)
for p in ["w1.xlsx", "w2.xlsx", "w3.xlsx"]:
    r = subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","lo",p],capture_output=True,timeout=180)
check("LibreOffice: w1 keeps print titles/cf/dv/zoom", (lambda b: b["Data"].print_title_rows.replace("$","") == "1:1"
      and sum(len(r.rules) for r in b["Data"].conditional_formatting) == 5 and len(b["Data"].data_validations.dataValidation) == 1
      and b["Data"].sheet_view.zoomScale == 120)(openpyxl.load_workbook("lo/w1.xlsx")))
b3 = openpyxl.load_workbook("lo/w3.xlsx")["Table 1"]
check("LibreOffice: w3 keeps freeze/filter/landscape-portrait", b3.freeze_panes == "A2" and b3.auto_filter.ref == "A1:B3"
      and b3.page_setup.orientation == "portrait", (b3.freeze_panes, b3.auto_filter.ref, b3.page_setup.orientation))

# ---------------------------------------------------------------- autofit leaves room for filter arrows
def widths(path):
    x = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode()
    return {int(m.group(1)): float(m.group(2)) for m in re.finditer(r'<col [^>]*?min="(\d+)"[^>]*?width="([\d.]+)"', x)}
def mk(name, title=False):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S"
    if title: ws.append(["Table 1: a long title sitting above the header row"]); ws.append([])
    ws.append(["mpg", "headroom"]); [ws.append([20 + i, 3]) for i in range(5)]
    wb.save(name); return os.path.abspath(name)
arrow = lambda n: round((n + 3) * 1.1 + 2, 2); plain = lambda n: max(6.0, round(n * 1.1 + 2, 2))
p = mk("af1.xlsx"); polish(p, autofit="1", autofilter="on"); w = widths(p)
check("autofit+autofilter same call: header room for arrow", w == {1: arrow(3), 2: arrow(8)}, w)
p = mk("af2.xlsx"); polish(p, autofilter="on"); polish(p, autofit="1"); w = widths(p)
check("autofit on sheet that already has a filter", w == {1: arrow(3), 2: arrow(8)}, w)
p = mk("af3.xlsx"); polish(p, autofit="1"); w = widths(p)
check("autofit without filter: no extra room", w == {1: plain(3), 2: plain(8)}, w)
p = mk("af4.xlsx"); polish(p, autofilter="on"); polish(p, autofit="1", autofilter="off"); w = widths(p)
check("autofit while removing filter: no extra room", w == {1: plain(3), 2: plain(8)}, w)
p = mk("af5.xlsx", title=True); polish(p, autofit="1", autofitmax="80", filterrange="A3:B8"); w = widths(p)
check("filterrange header in row 3: arrow room there, title row measured as is", w[2] == arrow(8) and w[1] == round(len("Table 1: a long title sitting above the header row") * 1.1 + 2, 2), w)
print("FAILURES:", fails)
