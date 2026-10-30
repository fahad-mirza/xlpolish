"""Notes, grouping, protection, unlocking, hidden sheets, dropdown references. Run via run_tests.sh."""
import os, re, zipfile, warnings, base64, hashlib, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.comments import Comment
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_core.py")).read().split("# ---------------------------------------------------------------- workbook 1")[0])
import lo

def build(name, notes_from_openpyxl=False):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Data"
    ws.append(["Economy", "Score", "Rank", "Region"])
    for i in range(1, 21): ws.append([f"Country {i}", 50 + i, i, "EAP" if i % 2 else "SAR"])
    s2 = wb.create_sheet("Notes sheet"); s2.append(["a", "b"]); s2["A1"].comment = Comment("pre-existing note", "Someone")
    s3 = wb.create_sheet("Third"); s3.append([1, 2, 3])
    wb.save(name); return os.path.abspath(name)

def ctypes(p): return zipfile.ZipFile(p).read("[Content_Types].xml").decode()
def part(p, n): return zipfile.ZipFile(p).read(n).decode()

# ================================================================ NOTES
f = build("n.xlsx")
rc, out = polish(f, sheet="Data", comment='B1 "Score: 0-100; higher is better"; A1 "Source: B-READY 2025\\nDraft"; D5 "R&D <check>"',
                 commentauthor="Fahad")
print("  ", out)
wb = openpyxl.load_workbook(f); d = wb["Data"]
check("3 notes added, text with ; & < > and line break", rc == 0 and d["B1"].comment.text == "Score: 0-100; higher is better"
      and d["A1"].comment.text == "Source: B-READY 2025\nDraft" and d["D5"].comment.text == "R&D <check>"
      and d["B1"].comment.author == "Fahad", (out, [(c, d[c].comment and d[c].comment.text) for c in ("A1","B1","D5")]))
check("other sheet's pre-existing note untouched", wb["Notes sheet"]["A1"].comment.text.startswith("pre-existing"))
vmls = [n for n in zipfile.ZipFile(f).namelist() if n.endswith(".vml")]
ids = [re.search(r'data="(\d+)"', part(f, n)).group(1) for n in vmls]
check("each notes drawing has its own id block", len(vmls) == 2 and len(set(ids)) == 2, (vmls, ids))
check("schema order", order_ok(f) == [], order_ok(f))
def lo_note_texts(path):
    """Note texts as LibreOffice imports them (via conversion to its own format)."""
    os.makedirs("lo_notes", exist_ok=True)
    subprocess.run(["soffice","--headless","--convert-to","ods","--outdir","lo_notes",path],capture_output=True,timeout=180)
    x = zipfile.ZipFile("lo_notes/" + os.path.basename(path).replace(".xlsx", ".ods")).read("content.xml").decode()
    out = []
    for a in re.findall(r"<office:annotation.*?</office:annotation>", x, flags=re.S):
        out.append("\n".join(re.sub(r"<[^>]+>", "", p) for p in re.findall(r"<text:p[^>]*>(.*?)</text:p>", a, flags=re.S))
                   .replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"))
    return sorted(out)
lo.stop()
t = lo_note_texts(f)
check("LibreOffice imports all notes with the right text", t == sorted(["Score: 0-100; higher is better", "Source: B-READY 2025\nDraft",
      "R&D <check>", "pre-existing note"]), t)

rc, out = polish(f, sheet="Data", comment='B1 "Updated text"; D5 ""; C2 "New one"'); print("  ", out)
d = openpyxl.load_workbook(f)["Data"]
check("second call: replace, delete, add", rc == 0 and d["B1"].comment.text == "Updated text" and d["D5"].comment is None
      and d["C2"].comment.text == "New one" and d["A1"].comment.text.startswith("Source"), out)
check("still exactly one notes part + drawing for Data", sum(1 for n in zipfile.ZipFile(f).namelist() if re.match(r"xl/comments", n)) == 2)

rc, out = polish(f, sheet="Notes sheet", comment='B1 "added next to the old one"', commentauthor="Fahad"); print("  ", out)
n2 = openpyxl.load_workbook(f)["Notes sheet"]
check("merge with a note written by another program", rc == 0 and n2["A1"].comment.text.startswith("pre-existing")
      and n2["A1"].comment.author == "Someone" and n2["B1"].comment.text == "added next to the old one", out)

rc, out = polish(f, sheet="Data", removecomments="1"); print("  ", out)
names = zipfile.ZipFile(f).namelist(); ct = ctypes(f); rels = part(f, "xl/worksheets/_rels/sheet1.xml.rels") if "xl/worksheets/_rels/sheet1.xml.rels" in names else ""
d = openpyxl.load_workbook(f)["Data"]
check("removecomments: notes, links and drawing gone", rc == 0 and all(d[c].comment is None for c in ("A1","B1","C2"))
      and "legacyDrawing" not in part(f, "xl/worksheets/sheet1.xml") and "comments" not in rels and "vmlDrawing" not in rels, rels)
dangling = [p for p in re.findall(r'PartName="/([^"]+)"', ct) if p not in names]
check("no content-type entries for deleted parts", dangling == [], dangling)
check("other sheet's notes still there", openpyxl.load_workbook(f)["Notes sheet"]["B1"].comment is not None)

# ================================================================ GROUPING
f = build("g.xlsx")
rc, out = polish(f, sheet="Data", grouprows="3:6 10:12", groupcols="B:C", collapse="1"); print("  ", out)
d = openpyxl.load_workbook(f)["Data"]
rd = d.row_dimensions
check("rows grouped and collapsed", rc == 0 and all(rd[r].outlineLevel == 1 and rd[r].hidden for r in (3,4,5,6,10,11,12))
      and not rd[7].hidden and rd[7].collapsed and rd[13].collapsed, [(r, rd[r].outlineLevel, rd[r].hidden, rd[r].collapsed) for r in (3,6,7,13)])
x = part(f, "xl/worksheets/sheet1.xml")
check("outline levels recorded in sheetFormatPr", 'outlineLevelRow="1"' in x and 'outlineLevelCol="1"' in x)
check("schema order", order_ok(f) == [], order_ok(f))
doc = lo.open_doc(f); sh = doc.Sheets.getByName("Data"); rows = sh.getRows(); cols = sh.getColumns()
check("LibreOffice: grouped rows/cols hidden, others visible", (not rows.getByIndex(3).IsVisible) and rows.getByIndex(6).IsVisible
      and (not cols.getByIndex(1).IsVisible) and cols.getByIndex(3).IsVisible); doc.close(True)

f = build("g2.xlsx")
rc, out = polish(f, sheet="Data", grouprows="3:12 5:7")
rd = openpyxl.load_workbook(f)["Data"].row_dimensions
check("nested groups get level 2", rc == 0 and rd[4].outlineLevel == 1 and rd[6].outlineLevel == 2 and 'outlineLevelRow="2"' in part(f, "xl/worksheets/sheet1.xml"))
rc, out = polish(f, sheet="Data", grouprows="30:40")   # beyond the data: rows get created
check("groups on empty rows beyond the data", rc == 0 and openpyxl.load_workbook(f)["Data"].row_dimensions[35].outlineLevel == 1 and order_ok(f) == [])
rc, out = polish(f, sheet="Data", ungroup="1"); print("  ", out)
x = part(f, "xl/worksheets/sheet1.xml")
check("ungroup removes everything", rc == 0 and "outlineLevel" not in x and "collapsed" not in x)
rc, out = polish(f, sheet="Data", grouprows="2:3 2:3 2:3 2:3 2:3 2:3 2:3 2:3")
check("more than 7 levels refused", rc == 198, out)

# ================================================================ PROTECTION
f = build("p.xlsx")
rc, out = polish(f, sheet="Data", protect="1", password="s3cret!", unlock="D2:D5 F:F 30:31", autofilter="on"); print("  ", out)
x = part(f, "xl/worksheets/sheet1.xml")
check("protection element with SHA-512 hash, no plain password", rc == 0 and 'algorithmName="SHA-512"' in x and "s3cret" not in x
      and 'spinCount="100000"' in x and 'autoFilter="0"' in x)
check("schema order", order_ok(f) == [], order_ok(f))
salt = base64.b64decode(re.search(r'saltValue="([^"]+)"', x).group(1))
h = hashlib.sha512(salt + "s3cret!".encode("utf-16-le")).digest()
for i in range(100000): h = hashlib.sha512(h + i.to_bytes(4, "little")).digest()
check("hash matches the documented Excel algorithm (computed independently)", base64.b64encode(h).decode() == re.search(r'hashValue="([^"]+)"', x).group(1))
doc = lo.open_doc(f); sh = doc.Sheets.getByName("Data")
locked = lambda a: sh.getCellRangeByName(a).CellProtection.IsLocked
check("LibreOffice: D2:D5, F100, row 30 unlocked; D6, A2 locked", (not locked("D2")) and (not locked("D5")) and locked("D6")
      and (not locked("F100")) and (not locked("H30")) and locked("A2"), [(a, locked(a)) for a in ("D2","D5","D6","F100","H30","A2")])
check("LibreOffice: sheet is protected", sh.isProtected())
try: sh.unprotect("wrong")
except Exception: pass
still = sh.isProtected()
sh.unprotect("s3cret!")
check("LibreOffice: wrong password refused, right password accepted", still and not sh.isProtected()); doc.close(True)

rc, out = polish(f, sheet="Data", unprotect="1"); print("  ", out)
check("unprotect", rc == 0 and "sheetProtection" not in part(f, "xl/worksheets/sheet1.xml"))
rc, out = polish(f, sheet="Third", protect="1"); doc = lo.open_doc(f); sh = doc.Sheets.getByName("Third")
p1 = sh.isProtected(); sh.unprotect(""); check("protect without password", rc == 0 and p1 and not sh.isProtected()); doc.close(True)
rc, out = polish(f, password="x", zoom="90"); check("password() without protect refused", rc == 198, out)
rc, out = polish(f, unlock="A1:ZZ99999"); check("huge bounded unlock range refused with advice", rc == 198 and "whole columns" in out, out)

# ================================================================ HIDE / SHOW
f = build("h.xlsx")
rc, out = polish(f, sheet="Notes sheet", hidesheet="1"); print("  ", out)
wb = openpyxl.load_workbook(f)
check("hide a non-active sheet", rc == 0 and wb["Notes sheet"].sheet_state == "hidden" and wb["Data"].sheet_state == "visible")
rc, out = polish(f, sheet="Data", hidesheet="1"); print("  ", out)
w = part(f, "xl/workbook.xml"); s3 = part(f, "xl/worksheets/sheet3.xml"); s1 = part(f, "xl/worksheets/sheet1.xml")
check("hiding the active sheet moves the active tab to the first visible one", rc == 0 and 'activeTab="2"' in w
      and 'tabSelected="1"' in s3 and 'tabSelected="1"' not in s1, re.findall(r"<workbookView[^>]*>", w))
doc = lo.open_doc(f); S = doc.Sheets
check("LibreOffice: Data and Notes sheet hidden, Third visible", (not S.getByName("Data").IsVisible)
      and (not S.getByName("Notes sheet").IsVisible) and S.getByName("Third").IsVisible); doc.close(True)
rc, out = polish(f, sheet="Third", hidesheet="1"); check("hiding the last visible sheet refused", rc == 198, out)
rc, out = polish(f, allsheets="1", hidesheet="1"); check("allsheets + hidesheet refused", rc == 198, out)
rc, out = polish(f, sheet="data", showsheet="1"); check("showsheet", rc == 0 and openpyxl.load_workbook(f)["Data"].sheet_state == "visible")

# ================================================================ everything together + round trip
f = build("all.xlsx")
rc, out = polish(f, sheet="Data", freeze="B2", autofilter="on", autofit="1", highlight="B2:B21 >= 60 lightgreen",
                 comment='B1 "Score"', grouprows="12:21", groupcols="C", unlock="D2:D21", protect="1", password="pw",
                 dropdown="D2:D21 EAP,SAR", landscape="1") if False else polish(f, sheet="Data", freeze="B2", autofilter="on", autofit="1",
                 highlight="B2:B21 >= 60 lightgreen", comment='B1 "Score"', grouprows="12:21", groupcols="C", unlock="D2:D21",
                 protect="1", password="pw", dropdown="D2:D21 EAP,SAR", orientation="landscape")
print("  ", out)
check("all features in one call", rc == 0 and order_ok(f) == [], (out, order_ok(f)))
os.makedirs("lo2", exist_ok=True)
subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","lo2",f],capture_output=True,timeout=180) if False else None
lo.stop()
r = subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","lo2",f],capture_output=True,timeout=180)
b = openpyxl.load_workbook("lo2/all.xlsx")["Data"]
check("LibreOffice round trip keeps note, groups, protection", b["B1"].comment is not None and b.row_dimensions[15].outlineLevel == 1
      and b.protection.sheet, (b["B1"].comment, b.row_dimensions[15].outlineLevel, b.protection.sheet))

# ================================================================ dropdown references made absolute
f = build("dd.xlsx")
for spec, want in [("D2:D9 =Third!A1:C1", "Third!$A$1:$C$1"), ("D2:D9 ='Notes sheet'!a1:a2", "'Notes sheet'!$A$1:$A$2"),
                   ("D2:D9 =Third!$A$1", "Third!$A$1"), ("D2:D9 =MyNamedList", "MyNamedList")]:
    rc, out = polish(f, sheet="Data", dropdown=spec)
    got = [d.formula1 for d in openpyxl.load_workbook(f)["Data"].data_validations.dataValidation][-1]
    check(f"dropdown source {spec.split(' ',1)[1]} -> {want}", rc == 0 and got == want, got)

# ================================================================ touching groups get a warning
def gfile(name):
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(list("ABCDEFGHIJ")); [ws.append(list(range(10))) for _ in range(20)]
    wb.save(name); return os.path.abspath(name)
f = gfile("t1.xlsx"); rc, out = polish(f, groupcols="C:E F:H")
check("touching column groups: warned once, at E/F, still applied", rc == 0 and out.count("note:") == 1
      and "columns E and F touch" in out and 'outlineLevel="1"' in part(f, "xl/worksheets/sheet1.xml"), out)
f = gfile("t2.xlsx"); rc, out = polish(f, groupcols="C:E G:I")
check("groups with a gap column: no warning", rc == 0 and "note:" not in out, out)
f = gfile("t3.xlsx"); rc, out = polish(f, grouprows="3:12 5:7")
check("nested groups: no warning", rc == 0 and "note:" not in out, out)
f = gfile("t4.xlsx"); rc, out = polish(f, grouprows="3:5 6:8 9:10")
check("three touching row groups: two warnings", rc == 0 and out.count("note:") == 2 and "rows 5 and 6" in out and "rows 8 and 9" in out, out)
f = gfile("t5.xlsx"); polish(f, grouprows="3:5"); rc, out = polish(f, grouprows="6:8")
check("touching a group from an earlier call: warned", rc == 0 and "rows 5 and 6 touch" in out, out)
f = gfile("t6.xlsx"); rc, out = polish(f, grouprows="3:12 3:5")
check("nested group sharing the first row: no warning", rc == 0 and "note:" not in out, out)

# ================================================================ summary side (button before or after groups)
f = gfile("s1.xlsx"); rc, out = polish(f, grouprows="5:8", collapse="1", summaryrows="above")
x = part(f, "xl/worksheets/sheet1.xml"); rd = openpyxl.load_workbook(f).active.row_dimensions
check("summaryabove: setting written, collapsed mark on the row before the group", rc == 0 and 'summaryBelow="0"' in x
      and rd[4].collapsed and not rd[9].collapsed and all(rd[r].hidden for r in (5,6,7,8)) and not rd[4].hidden, (out, x[:300]))
check("summaryabove: openpyxl reads it", openpyxl.load_workbook(f).active.sheet_properties.outlinePr.summaryBelow is False)
f = gfile("s2.xlsx"); rc, out = polish(f, groupcols="C:E", collapse="1", summarycols="left")
x = part(f, "xl/worksheets/sheet1.xml")
check("summaryleft: setting written, collapsed mark on column B", rc == 0 and 'summaryRight="0"' in x
      and re.search(r'<col [^>]*collapsed="1"[^>]*max="2" min="2"', x) and not re.search(r'<col [^>]*collapsed="1"[^>]*min="6"', x), x[:600])
rc, out = polish(f, summaryrows="below", summarycols="right")
check("summarybelow/right restore Excel's default (setting removed)", rc == 0 and "outlinePr" not in part(f, "xl/worksheets/sheet1.xml"), out)
f = gfile("s3.xlsx"); rc, out = polish(f, tabcolor="navy", fitwidth="1", summaryrows="above", summarycols="left", grouprows="3:4")
x = part(f, "xl/worksheets/sheet1.xml"); sp = re.search(r"<sheetPr>(.*?)</sheetPr>", x).group(1)
kids = re.findall(r"<(\w+)", sp)
check("sheet properties stay in Excel's order (tab colour, outline, page setup)", rc == 0 and kids == ["tabColor", "outlinePr", "pageSetUpPr"] and order_ok(f) == [], kids)
lo.stop()
os.makedirs("lo_sum", exist_ok=True)
subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","lo_sum",gfile("s4.xlsx")],capture_output=True,timeout=180)
f = gfile("s4.xlsx"); polish(f, grouprows="5:8", collapse="1", summaryrows="above")
doc = lo.open_doc(f); sh = doc.Sheets.getByIndex(0); rows = sh.getRows()
check("LibreOffice: grouped rows 5-8 hidden, row 4 visible", (not rows.getByIndex(5).IsVisible) and rows.getByIndex(3).IsVisible and rows.getByIndex(8).IsVisible)
doc.close(True); lo.stop()

# ================================================================ 1.3.0: filter arrows, validation, prompts, icons, ranks, header/footer, hide
def tbl(name):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "T"
    ws.append(["Economy", "Score", "Rank", "Year", "", "Note"])
    for i in range(1, 21): ws.append([f"Economy {i}", 40 + i * 2.5, i, 2020 + i % 5, None, "ok" if i % 3 else None])
    wb.save(name); return os.path.abspath(name)
def sheet_xml(f): return part(f, "xl/worksheets/sheet1.xml")

# --- nofilterarrow
f = tbl("fa.xlsx"); rc, out = polish(f, autofilter="on", autofit="1", nofilterarrow="E")
ws = openpyxl.load_workbook(f).active; fcs = ws.auto_filter.filterColumn
check("nofilterarrow: column E arrow hidden, filter still A1:F21", rc == 0 and ws.auto_filter.ref == "A1:F21"
      and len(fcs) == 1 and fcs[0].colId == 4 and fcs[0].hiddenButton, (out, ws.auto_filter.ref, [(c.colId, c.hiddenButton) for c in fcs]))
check("nofilterarrow: schema order", order_ok(f) == [], order_ok(f))
rc, out = polish(f, nofilterarrow="F")
fcs = openpyxl.load_workbook(f).active.auto_filter.filterColumn
check("nofilterarrow on an existing filter adds F, keeps E, in column order", rc == 0 and [c.colId for c in fcs] == [4, 5], [c.colId for c in fcs])
rc, out = polish(tbl("fa2.xlsx"), nofilterarrow="E"); check("nofilterarrow without a filter refused", rc == 198 and "needs a filter" in out, out)
rc, out = polish(tbl("fa3.xlsx"), filterrange="A1:C21", nofilterarrow="E"); check("nofilterarrow outside the filter range refused", rc == 198 and "outside" in out, out)
rc, out = polish(tbl("fa4.xlsx"), removefilter="1", autofilter="off", nofilterarrow="E") if False else polish(tbl("fa4.xlsx"), autofilter="off", nofilterarrow="E")
check("nofilterarrow with removefilter refused", rc == 198, out)

# --- validate
f = tbl("va.xlsx")
rc, out = polish(f, validate='B2:B21 decimal between 0 100; C2:C21 whole >= 1; D2:D21 date between 2020-01-01 2026-12-31; F2:F21 length <= 50 "Keep notes short"')
dv = {str(d.sqref): d for d in openpyxl.load_workbook(f).active.data_validations.dataValidation}
check("validate: four rules written", rc == 0 and len(dv) == 4, (out, list(dv)))
b = dv["B2:B21"]; c = dv["C2:C21"]; dd = dv["D2:D21"]; ff = dv["F2:F21"]
check("validate: decimal between 0 and 100, plain default message", b.type == "decimal" and b.operator in (None, "between")
      and b.formula1 == "0" and b.formula2 == "100" and b.error == "Enter a number from 0 to 100.", (b.type, b.operator, b.formula1, b.formula2, b.error))
check("validate: whole >= 1", c.type == "whole" and c.operator == "greaterThanOrEqual" and c.formula1 == "1")
check("validate: dates converted to Excel day numbers (2020-01-01 = 43831)", dd.type == "date" and dd.formula1 == "43831" and dd.formula2 == "46387", (dd.formula1, dd.formula2))
check("validate: text length with custom message", ff.type == "textLength" and ff.operator == "lessThanOrEqual" and ff.error == "Keep notes short")
for bad, why in [("B2:B9 decimal between 9 1", "bounds reversed"), ("B2:B9 number > 1", "unknown type"), ("B2:B9 date > 31/12/2026", "bad date"),
                 ("B2:B9 whole > 1.5", "fraction for whole"), ("B2:B9 decimal between 1", "missing bound"), ("B3:B4 decimal > 0", "overlaps existing rule")]:
    rc, out = polish(f, validate=bad); check(f"validate refused: {why}", rc == 198, out)
doc = lo.open_doc(f); sh = doc.Sheets.getByIndex(0); v = sh.getCellRangeByName("B5").Validation
from com.sun.star.sheet.ValidationType import DECIMAL
check("LibreOffice: B5 has the decimal validation with the right bounds", v.Type == DECIMAL and v.getFormula1() == "0" and v.getFormula2() == "100", (v.Type, v.getFormula1(), v.getFormula2()))
doc.close(True); lo.stop()

# --- prompt
f = tbl("pr.xlsx")
rc, out = polish(f, dropdown="E2:E21 Yes,No", prompt='E2:E21 "Was this economy reviewed?"; A2:A21 "Economy names are fixed"')
dvs = {str(d.sqref): d for d in openpyxl.load_workbook(f).active.data_validations.dataValidation}
check("prompt joins the dropdown on the same cells (one rule, list + prompt)", rc == 0 and len(dvs) == 2
      and dvs["E2:E21"].type == "list" and dvs["E2:E21"].prompt == "Was this economy reviewed?" and dvs["E2:E21"].showInputMessage, (out, list(dvs)))
check("prompt on its own cells: a rule that accepts anything", dvs["A2:A21"].type in (None, "none") and dvs["A2:A21"].prompt.startswith("Economy"))
rc, out = polish(f, dropdown="E2:E21 Yes,No,Unsure")
e = [d for d in openpyxl.load_workbook(f).active.data_validations.dataValidation if str(d.sqref) == "E2:E21"]
check("re-running a dropdown replaces it and keeps its prompt", rc == 0 and len(e) == 1 and e[0].formula1 == '"Yes,No,Unsure"' and e[0].prompt.startswith("Was"), [(d.formula1, d.prompt) for d in e])
rc, out = polish(f, prompt='E5:E30 "partial"'); check("prompt partly overlapping another rule refused", rc == 198 and "overlaps" in out, out)
check("validation schema order", order_ok(f) == [], order_ok(f))

# --- icon sets, top/bottom, average
f = tbl("ic.xlsx")
rc, out = polish(f, iconset="B2:B21 traffic; C2:C21 arrows5 percentile reverse iconsonly", top="B2:B21 3 lightgreen", bottom="B2:B21 10% pink darkred",
                 average="C2:C21 below FFC7CE")
ws = openpyxl.load_workbook(f).active
rules = [(str(r.sqref), x) for r in ws.conditional_formatting for x in r.rules]
kinds = sorted(x.type for _, x in rules)
check("icon sets, top, bottom and average rules written", rc == 0 and kinds == ["aboveAverage", "iconSet", "iconSet", "top10", "top10"], (out, kinds))
icons = {sq: x.iconSet for sq, x in rules if x.type == "iconSet"}
t = icons["B2:B21"]; a5 = icons["C2:C21"]
check("traffic lights: 3 thresholds at 0/33/67 percent", t.iconSet == "3TrafficLights1" and [c.val for c in t.cfvo] == [0, 33, 67] and t.cfvo[0].type == "percent",
      (t.iconSet, [(c.type, c.val) for c in t.cfvo]))
check("arrows5 percentile reverse iconsonly", a5.iconSet == "5Arrows" and a5.reverse and a5.showValue is False and [c.val for c in a5.cfvo] == [0, 20, 40, 60, 80]
      and a5.cfvo[1].type == "percentile")
tops = [x for _, x in rules if x.type == "top10"]
check("top 3 and bottom 10%", sorted((x.rank, bool(x.percent), bool(x.bottom)) for x in tops) == [(3, False, False), (10, True, True)])
avg = [x for _, x in rules if x.type == "aboveAverage"][0]
check("below average", avg.aboveAverage is False and avg.dxfId is not None)
pr = [x.priority for _, x in rules]; check("all priorities unique", len(set(pr)) == len(pr), pr)
for bad, why in [("B2:B9 lights", "unknown icon style"), ("B2:B9 traffic upside", "unknown extra word")]:
    rc, out = polish(f, iconset=bad); check(f"iconset refused: {why}", rc == 198, out)
rc, out = polish(f, top="B2:B9 0 red"); check("top() N of 0 refused", rc == 198, out)
rc, out = polish(f, average="B2:B9 middle red"); check("average() needs above/below", rc == 198, out)

# --- header / footer
f = tbl("hf.xlsx")
rc, out = polish(f, header='"Draft & confidential"', footer='"[file] | Page [page] of [pages] | [date]"')
x = sheet_xml(f)
check("header centred with & escaped; footer in three sections with codes", rc == 0
      and "<oddHeader>&amp;CDraft &amp;&amp; confidential</oddHeader>" in x
      and "<oddFooter>&amp;L&amp;F&amp;CPage &amp;P of &amp;N&amp;R&amp;D</oddFooter>" in x, re.findall(r"<headerFooter>.*?</headerFooter>", x))
check("header/footer schema order", order_ok(f) == [], order_ok(f))
wsx = openpyxl.load_workbook(f).active
check("openpyxl reads footer sections", wsx.oddFooter.center.text == "Page &P of &N" and wsx.oddFooter.left.text == "&F", (wsx.oddFooter.left.text, wsx.oddFooter.center.text))
rc, out = polish(f, footer='""'); x = sheet_xml(f)
check("footer(\"\") removes the footer, keeps the header", rc == 0 and "oddFooter" not in x and "oddHeader" in x)
rc, out = polish(f, header='""'); check("removing both drops the element", rc == 0 and "headerFooter" not in sheet_xml(f))
rc, out = polish(f, footer='"Page [pg]"'); check("unknown placeholder refused", rc == 198 and "[page]" in out, out)
rc, out = polish(f, footer='"a | b"'); check("two sections refused", rc == 198, out)

# --- hide / show
f = tbl("hd.xlsx")
rc, out = polish(f, hiderows="5:7", hidecols="E F")
rd = openpyxl.load_workbook(f).active
check("rows 5-7 and columns E, F hidden", rc == 0 and all(rd.row_dimensions[r].hidden for r in (5,6,7)) and not rd.row_dimensions[8].hidden
      and rd.column_dimensions["E"].hidden and rd.column_dimensions["F"].hidden)
doc = lo.open_doc(f); sh = doc.Sheets.getByIndex(0)
check("LibreOffice: row 6 and column E hidden, row 8 and column D visible", (not sh.getRows().getByIndex(5).IsVisible) and sh.getRows().getByIndex(7).IsVisible
      and (not sh.getColumns().getByIndex(4).IsVisible) and sh.getColumns().getByIndex(3).IsVisible)
doc.close(True); lo.stop()
rc, out = polish(f, showrows="5:7", showcols="E:F")
rd = openpyxl.load_workbook(f).active
check("showrows/showcols bring them back", rc == 0 and not any(rd.row_dimensions[r].hidden for r in (5,6,7)) and not rd.column_dimensions["E"].hidden)
check("hide/show schema order", order_ok(f) == [], order_ok(f))

# --- all of it in one call, then LibreOffice round trip
f = tbl("all13.xlsx")
rc, out = polish(f, autofilter="on", nofilterarrow="E", validate="B2:B21 decimal between 0 100", prompt='B2:B21 "Score from 0 to 100"',
                 iconset="C2:C21 flags", top="B2:B21 5 lightgreen", average="B2:B21 below pink", header='"xlpolish test"',
                 footer='"Page [page] of [pages]"', hidecols="F", freeze="B2", autofit="1")
check("all 1.3.0 options together", rc == 0 and order_ok(f) == [], (out, order_ok(f)))
os.makedirs("lo13", exist_ok=True)
subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","lo13",f],capture_output=True,timeout=180)
b = openpyxl.load_workbook("lo13/all13.xlsx").active
check("LibreOffice round trip keeps validation, prompt, icon set and footer",
      any(d.type == "decimal" and d.prompt for d in b.data_validations.dataValidation)
      and any(x.type == "iconSet" for r in b.conditional_formatting for x in r.rules) and "Page" in (b.oddFooter.center.text or ""),
      ([(d.type, d.prompt) for d in b.data_validations.dataValidation], b.oddFooter.center.text))

# ================================================================ rule priority: newer rules on top, like Excel
f = tbl("pri.xlsx")
rc, out = polish(f, colorscale="B2:B21", databar="C2:C21", highlight="B2:B21 >= 80 lightgreen")
pri = {x.type: x.priority for r in openpyxl.load_workbook(f).active.conditional_formatting for x in r.rules}
check("within one call: highlight above colour scale and data bar", rc == 0 and pri["expression"] == 1 and sorted(pri.values()) == [1, 2, 3], pri)
rc, out = polish(f, top="C2:C21 3 lightgreen", average="B2:B21 below pink")
rl = [(x.type, x.priority) for r in openpyxl.load_workbook(f).active.conditional_formatting for x in r.rules]
p = dict(rl)
check("second call's rules go to the top, earlier rules move down", rc == 0 and p["top10"] == 1 and p["aboveAverage"] == 2
      and p["expression"] == 3 and sorted(v for _, v in rl) == [1, 2, 3, 4, 5], rl)
print("FAILURES:", fails)
