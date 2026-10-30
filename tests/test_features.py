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
print("FAILURES:", fails)
