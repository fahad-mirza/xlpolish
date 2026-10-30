package xlpolish;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * Applies an xlpolish Spec to a workbook. All options are parsed and validated
 * before anything is changed, and the file is only rewritten at the very end, so
 * an error at any point leaves the workbook untouched.
 */
final class Engine {

    // ------------------------------------------------------ parsed options

    private static final class ColorScale { String sqref; String[] colors; }
    private static final class DataBar { String sqref; String color; }
    private static final class Highlight { String sqref; String anchor; String op; String formula; String fill; String font; }
    private static final class Dropdown { String sqref; String formula; }
    private static final class ColWidth { int c1, c2; double w; }

    private final Spec spec;
    private int freezeRows = -1, freezeCols = -1;
    private String freezeCell;
    private String tabArgb;                 // null = leave, "" = remove
    private int[] filterRect;
    private final List<ColWidth> widths = new ArrayList<>();
    private String titleRows, titleCols;    // "$1:$2", "$A:$A"
    private final List<ColorScale> scales = new ArrayList<>();
    private final List<DataBar> bars = new ArrayList<>();
    private final List<Highlight> highlights = new ArrayList<>();
    private final List<Dropdown> dropdowns = new ArrayList<>();
    private final TreeMap<Long, String> notes = new TreeMap<>();
    private final List<int[]> groupRows = new ArrayList<>();
    private final List<int[]> groupCols = new ArrayList<>();
    private final List<int[]> unlockRects = new ArrayList<>();
    /** Warnings to show under a sheet's summary line (the call still succeeds). */
    private final List<String> notesOut = new ArrayList<>();

    Engine(Spec spec) throws XLPolishException {
        this.spec = spec;
        if (!spec.hasAction()) throw XLPolishException.syntax("nothing to do; specify at least one option such as freeze() or autofit");
        prepare();
    }

    // ------------------------------------------------------------- parsing

    private void prepare() throws XLPolishException {
        if (spec.freeze != null) {
            int[] rc = Cells.parseCell(spec.freeze);
            if (rc == null) throw XLPolishException.syntax("freeze() must be a cell such as A2 or B2; got \"" + spec.freeze + "\"");
            freezeRows = rc[0] - 1;
            freezeCols = rc[1] - 1;
            freezeCell = Cells.ref(rc[0], rc[1]);
        }
        if (spec.tabcolor != null) tabArgb = "none".equalsIgnoreCase(spec.tabcolor) ? "" : Colors.argb(spec.tabcolor);
        if (spec.filterrange != null) {
            filterRect = Cells.parseRect(spec.filterrange);
            if (filterRect[0] == filterRect[2]) throw XLPolishException.syntax("filterrange() needs a header row and at least one data row");
        }
        if (spec.colwidth != null) parseColWidths(spec.colwidth);
        if (spec.orientation != null && !spec.orientation.equals("landscape") && !spec.orientation.equals("portrait")) {
            throw XLPolishException.syntax("internal: orientation must be landscape or portrait");
        }
        if (spec.printtitles != null) parsePrintTitles(spec.printtitles);
        for (String r : spec.colorscale) scales.add(parseColorScale(r));
        for (String r : spec.databar) bars.add(parseDataBar(r));
        for (String r : spec.highlight) highlights.add(parseHighlight(r));
        for (String r : spec.dropdown) dropdowns.add(parseDropdown(r));
        for (String r : spec.comment) parseNote(r);
        if (spec.groupRows != null) parseRowGroups(spec.groupRows);
        if (spec.groupCols != null) parseColGroups(spec.groupCols);
        if (spec.collapse && groupRows.isEmpty() && groupCols.isEmpty()) {
            throw XLPolishException.syntax("collapse requires grouprows() or groupcols()");
        }
        if (spec.ungroup && (!groupRows.isEmpty() || !groupCols.isEmpty())) {
            throw XLPolishException.syntax("ungroup may not be combined with grouprows() or groupcols()");
        }
        if (spec.password != null && !Boolean.TRUE.equals(spec.protect)) {
            throw XLPolishException.syntax("password() requires protect");
        }
        for (String r : spec.unlock) {
            for (String piece : r.trim().split("[\\s,]+")) if (!piece.isEmpty()) unlockRects.add(Cells.parseRect(piece));
        }
        if (Boolean.TRUE.equals(spec.hideSheet) && spec.mode == Spec.Mode.ALL) {
            throw XLPolishException.syntax("hidesheet cannot be combined with allsheets (a workbook needs at least one visible sheet)");
        }
    }

    private void parseNote(String r) throws XLPolishException {
        String s = r.trim();
        int sp = -1;
        for (int i = 0; i < s.length(); i++) if (Character.isWhitespace(s.charAt(i))) { sp = i; break; }
        String cell = sp < 0 ? s : s.substring(0, sp);
        String text = sp < 0 ? "" : s.substring(sp + 1).trim();
        int[] rc = Cells.parseCell(cell);
        if (rc == null) throw XLPolishException.syntax("comment() takes a single cell and then the text, e.g. comment(B1 \"Source: B-READY\"); got: " + r);
        if (text.length() >= 2 && text.startsWith("\"") && text.endsWith("\"")) text = text.substring(1, text.length() - 1);
        text = text.replace("\\n", "\n");
        if (text.length() > 32767) throw XLPolishException.syntax("comment() text is longer than Excel's 32,767-character limit");
        notes.put(Notes.key(rc[0], rc[1]), text);
    }

    private void parseRowGroups(String s) throws XLPolishException {
        for (String tok : s.trim().split("[\\s,]+")) {
            if (tok.isEmpty()) continue;
            String u = tok.replace("$", "");
            if (!u.matches("[0-9]+(:[0-9]+)?")) throw XLPolishException.syntax("grouprows() takes rows such as 5:20 or 7; got \"" + tok + "\"");
            String[] p = u.split(":");
            long a = Long.parseLong(p[0]), b = Long.parseLong(p[p.length - 1]);
            if (a < 1 || a > b || b > Cells.MAX_ROWS) throw XLPolishException.syntax("invalid rows in grouprows(): " + tok);
            groupRows.add(new int[] { (int) a, (int) b });
        }
        if (groupRows.isEmpty()) throw XLPolishException.syntax("grouprows() is empty");
    }

    private void parseColGroups(String s) throws XLPolishException {
        for (String tok : s.trim().split("[\\s,]+")) {
            if (tok.isEmpty()) continue;
            String u = tok.replace("$", "").toUpperCase(Locale.ROOT);
            if (!u.matches("[A-Z]{1,3}(:[A-Z]{1,3})?")) throw XLPolishException.syntax("groupcols() takes columns such as C:F or H; got \"" + tok + "\"");
            String[] p = u.split(":");
            int a = Cells.colNum(p[0]), b = Cells.colNum(p[p.length - 1]);
            if (a > b || b > Cells.MAX_COLS) throw XLPolishException.syntax("invalid columns in groupcols(): " + tok);
            groupCols.add(new int[] { a, b });
        }
        if (groupCols.isEmpty()) throw XLPolishException.syntax("groupcols() is empty");
    }

    private static final Pattern CW = Pattern.compile("^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?=([0-9]*\\.?[0-9]+)$");

    private void parseColWidths(String s) throws XLPolishException {
        for (String tok : Spec.tokens(s)) {
            Matcher m = CW.matcher(tok);
            if (!m.matches()) throw XLPolishException.syntax("colwidth() items look like A=30 or B:D=12; got \"" + tok + "\"");
            ColWidth cw = new ColWidth();
            cw.c1 = Cells.colNum(m.group(1).toUpperCase(Locale.ROOT));
            cw.c2 = m.group(2) == null ? cw.c1 : Cells.colNum(m.group(2).toUpperCase(Locale.ROOT));
            cw.w = Double.parseDouble(m.group(3));
            if (cw.c1 > cw.c2 || cw.c2 > Cells.MAX_COLS) throw XLPolishException.syntax("invalid columns in colwidth(): " + tok);
            if (cw.w > 255) throw XLPolishException.syntax("column width must be 0-255 characters: " + tok);
            widths.add(cw);
        }
        if (widths.isEmpty()) throw XLPolishException.syntax("colwidth() is empty");
    }

    private void parsePrintTitles(String s) throws XLPolishException {
        for (String tok : Spec.tokens(s)) {
            String u = tok.toUpperCase(Locale.ROOT).replace("$", "");
            if (u.matches("[0-9]+(:[0-9]+)?")) {
                String[] p = u.split(":");
                int a = Integer.parseInt(p[0]), b = Integer.parseInt(p[p.length - 1]);
                if (a < 1 || a > b || b > Cells.MAX_ROWS) throw XLPolishException.syntax("invalid rows in printtitles(): " + tok);
                titleRows = "$" + a + ":$" + b;
            } else if (u.matches("[A-Z]{1,3}(:[A-Z]{1,3})?")) {
                String[] p = u.split(":");
                int a = Cells.colNum(p[0]), b = Cells.colNum(p[p.length - 1]);
                if (a > b || b > Cells.MAX_COLS) throw XLPolishException.syntax("invalid columns in printtitles(): " + tok);
                titleCols = "$" + Cells.colName(a) + ":$" + Cells.colName(b);
            } else {
                throw XLPolishException.syntax("printtitles() takes rows such as 1 or 1:2 and/or columns such as A or A:B; got \"" + tok + "\"");
            }
        }
    }

    private ColorScale parseColorScale(String r) throws XLPolishException {
        List<String> t = Spec.tokens(r);
        if (t.isEmpty()) throw XLPolishException.syntax("colorscale() needs a range");
        ColorScale cs = new ColorScale();
        cs.sqref = Cells.sqref(t.get(0));
        List<String> cols = t.subList(1, t.size());
        if (cols.isEmpty()) {
            cs.colors = new String[] { "FFF8696B", "FFFFEB84", "FF63BE7B" };     // red - yellow - green
        } else if (cols.size() == 2 || cols.size() == 3) {
            cs.colors = new String[cols.size()];
            for (int i = 0; i < cols.size(); i++) cs.colors[i] = Colors.argb(cols.get(i));
        } else {
            throw XLPolishException.syntax("colorscale() takes a range and then 2 or 3 colours (low [middle] high): " + r);
        }
        return cs;
    }

    private DataBar parseDataBar(String r) throws XLPolishException {
        List<String> t = Spec.tokens(r);
        if (t.isEmpty() || t.size() > 2) throw XLPolishException.syntax("databar() takes a range and an optional colour: " + r);
        DataBar d = new DataBar();
        d.sqref = Cells.sqref(t.get(0));
        d.color = t.size() == 2 ? Colors.argb(t.get(1)) : "FF638EC6";
        return d;
    }

    private static final Map<String, String> OPS = new HashMap<>();
    static {
        // Stata-style operator -> Excel formula operator
        OPS.put("==", "="); OPS.put("=", "=");
        OPS.put("!=", "<>"); OPS.put("~=", "<>"); OPS.put("<>", "<>");
        OPS.put(">", ">"); OPS.put(">=", ">=");
        OPS.put("<", "<"); OPS.put("<=", "<=");
    }

    private Highlight parseHighlight(String r) throws XLPolishException {
        List<String> t = Spec.tokens(r);
        if (t.size() < 4 || t.size() > 5) {
            throw XLPolishException.syntax("highlight() takes: range operator value fill-colour [text-colour], "
                    + "e.g. highlight(B2:H164 == 5 lightgreen); got: " + r);
        }
        Highlight h = new Highlight();
        h.sqref = Cells.sqref(t.get(0));
        int[] first = Cells.parseRect(h.sqref.split(" ")[0]);
        h.anchor = Cells.ref(first[0], first[1]);
        h.op = OPS.get(t.get(1));
        if (h.op == null) throw XLPolishException.syntax("highlight() operator must be one of == != > >= < <=; got \"" + t.get(1) + "\"");
        String v = t.get(2);
        if (v.matches("[-+]?([0-9]+\\.?[0-9]*|\\.[0-9]+)([eE][-+]?[0-9]+)?")) {
            h.formula = v;
        } else {
            h.formula = "\"" + v.replace("\"", "\"\"") + "\"";
            if (h.formula.length() > 255) throw XLPolishException.syntax("highlight() text value is too long");
        }
        h.fill = Colors.argb(t.get(3));
        h.font = t.size() == 5 ? Colors.argb(t.get(4)) : null;
        return h;
    }

    private Dropdown parseDropdown(String r) throws XLPolishException {
        String s = r.trim();
        int sp = -1;
        for (int i = 0; i < s.length(); i++) if (Character.isWhitespace(s.charAt(i))) { sp = i; break; }
        if (sp < 0) throw XLPolishException.syntax("dropdown() takes a range and then the choices, e.g. dropdown(C2:C100 Yes,No): " + r);
        Dropdown d = new Dropdown();
        d.sqref = Cells.sqref(s.substring(0, sp));
        String rest = s.substring(sp + 1).trim();
        if (rest.length() >= 2 && rest.startsWith("\"") && rest.endsWith("\"")) rest = rest.substring(1, rest.length() - 1).trim();
        if (rest.isEmpty()) throw XLPolishException.syntax("dropdown() has no choices: " + r);
        if (rest.startsWith("=")) {
            d.formula = absoluteRef(rest.substring(1).trim());          // e.g. =Lists!A1:A5 -> Lists!$A$1:$A$5
        } else {
            List<String> items = new ArrayList<>();
            for (String it : rest.split(",")) if (!it.trim().isEmpty()) items.add(it.trim());
            if (items.isEmpty()) throw XLPolishException.syntax("dropdown() has no choices: " + r);
            d.formula = "\"" + String.join(",", items) + "\"";
            if (d.formula.length() > 257) {
                throw XLPolishException.syntax("dropdown() list is longer than Excel's 255-character limit; "
                        + "put the choices in cells and use dropdown(range =Sheet!A1:A20)");
            }
        }
        return d;
    }

    private static final Pattern REF = Pattern.compile(
            "^((?:'(?:[^']|'')+'|[^!'\\s]+)!)?\\$?([A-Za-z]{1,3})\\$?([0-9]+)(?::\\$?([A-Za-z]{1,3})\\$?([0-9]+))?$");

    /**
     * Makes a plain cell reference absolute ($A$1). A dropdown's source must not
     * move from cell to cell, and typing $ in Stata would start a global macro,
     * so users can write Lists!A1:A5. Anything else (a named range) is left as is.
     */
    static String absoluteRef(String ref) {
        Matcher m = REF.matcher(ref);
        if (!m.matches()) return ref;
        String sheet = m.group(1) == null ? "" : m.group(1);
        String out = sheet + "$" + m.group(2).toUpperCase(Locale.ROOT) + "$" + m.group(3);
        if (m.group(4) != null) out += ":$" + m.group(4).toUpperCase(Locale.ROOT) + "$" + m.group(5);
        return out;
    }

    // -------------------------------------------------------------- apply

    List<String> run() throws XLPolishException {
        List<String> out = new ArrayList<>();
        try (Pkg pkg = new Pkg(Paths.get(spec.file))) {
            Book book = new Book(pkg);
            List<Book.Sheet> targets = book.targets(spec);
            for (Book.Sheet sh : targets) {
                List<String> done = applySheet(book, sh);
                if (spec.hideSheet != null) done.add(setVisible(book, sh, !spec.hideSheet));
                out.add("Sheet \"" + sh.name + "\": " + (done.isEmpty() ? "no changes needed" : String.join("; ", done)));
                for (String n : notesOut) out.add("  " + n);
                notesOut.clear();
            }
            if (Boolean.TRUE.equals(spec.hideSheet)) fixActiveTab(book);
            pkg.save();
        }
        return out;
    }

    private List<String> applySheet(Book book, Book.Sheet sh) throws XLPolishException {
        Element ws = book.pkg.doc(sh.part).getDocumentElement();
        List<String> done = new ArrayList<>();
        boolean changed = false;

        if (freezeCell != null) { changed |= freeze(ws); done.add(describeFreeze()); }
        if (spec.gridlines != null) { gridlines(ws, spec.gridlines); changed = true; done.add(spec.gridlines ? "gridlines shown" : "gridlines hidden"); }
        if (spec.zoom != null) { zoom(ws, spec.zoom); changed = true; done.add("zoom " + spec.zoom + "%"); }
        if (tabArgb != null) { tabColor(ws, tabArgb); changed = true; done.add(tabArgb.isEmpty() ? "tab colour removed" : "tab colour set"); }
        if (spec.autofit) {
            int n = autofit(book, ws);
            if (n > 0) { changed = true; done.add("auto-fit " + n + (n == 1 ? " column" : " columns")); }
            else done.add("auto-fit: sheet has no data");
        }
        if (!widths.isEmpty()) {
            Element cols = Pkg.ensure(ws, "cols", Pkg.WORKSHEET_ORDER);
            for (ColWidth cw : widths) for (int c = cw.c1; c <= cw.c2; c++) setWidth(cols, c, cw.w);
            changed = true;
            done.add("column widths set");
        }
        if (spec.autofilter != null) {
            String msg = autofilter(book, sh, ws, spec.autofilter);
            changed = true;
            done.add(msg);
        }
        if (spec.orientation != null || spec.fitwidth) {
            pageSetup(ws);
            changed = true;
            if (spec.orientation != null) done.add(spec.orientation);
            if (spec.fitwidth) done.add("fit to 1 page wide");
        }
        if (titleRows != null || titleCols != null) {
            String ref = "";
            String q = quoteSheet(sh.name);
            if (titleCols != null) ref = q + "!" + titleCols;
            if (titleRows != null) ref = ref.isEmpty() ? q + "!" + titleRows : ref + "," + q + "!" + titleRows;
            setDefinedName(book, "_xlnm.Print_Titles", sh.index, ref, false);
            done.add("print titles " + (titleRows != null ? "rows " + titleRows.replace("$", "") : "")
                    + (titleRows != null && titleCols != null ? ", " : "")
                    + (titleCols != null ? "columns " + titleCols.replace("$", "") : ""));
        }
        if (!scales.isEmpty() || !bars.isEmpty() || !highlights.isEmpty()) {
            int pr = maxPriority(ws);
            for (ColorScale cs : scales) addColorScale(ws, cs, ++pr);
            for (DataBar d : bars) addDataBar(ws, d, ++pr);
            for (Highlight h : highlights) addHighlight(book, ws, h, ++pr);
            changed = true;
            int n = scales.size() + bars.size() + highlights.size();
            done.add(n + " conditional format" + (n == 1 ? "" : "s") + " added");
        }
        if (!dropdowns.isEmpty()) {
            for (Dropdown d : dropdowns) addDropdown(ws, d);
            changed = true;
            done.add(dropdowns.size() + " dropdown list" + (dropdowns.size() == 1 ? "" : "s") + " added");
        }
        if (!notes.isEmpty() || spec.removeComments) {
            int[] n = Notes.apply(book, sh, ws, notes, spec.commentAuthor, spec.removeComments);
            changed = true;
            if (spec.removeComments) done.add("all notes removed");
            if (n[0] > 0) done.add(n[0] + " note" + (n[0] == 1 ? "" : "s") + " added");
            if (n[1] > 0) done.add(n[1] + " note" + (n[1] == 1 ? "" : "s") + " deleted");
        }
        if (spec.summaryAbove != null || spec.summaryLeft != null) {
            Outline.summarySide(ws, spec.summaryAbove, spec.summaryLeft);
            changed = true;
            if (spec.summaryAbove != null) done.add(spec.summaryAbove ? "summary rows above detail" : "summary rows below detail");
            if (spec.summaryLeft != null) done.add(spec.summaryLeft ? "summary columns left of detail" : "summary columns right of detail");
        }
        if (spec.ungroup) {
            changed |= Outline.ungroup(ws);
            done.add("grouping removed");
        }
        if (!groupRows.isEmpty()) {
            notesOut.addAll(Outline.groupRows(ws, groupRows, spec.collapse));
            changed = true;
            done.add(groupRows.size() + " row group" + (groupRows.size() == 1 ? "" : "s") + (spec.collapse ? " (collapsed)" : ""));
        }
        if (!groupCols.isEmpty()) {
            notesOut.addAll(Outline.groupCols(ws, groupCols, spec.collapse));
            changed = true;
            done.add(groupCols.size() + " column group" + (groupCols.size() == 1 ? "" : "s") + (spec.collapse ? " (collapsed)" : ""));
        }
        if (!unlockRects.isEmpty()) {
            Protect.unlock(book, ws, unlockRects);
            changed = true;
            done.add("cells unlocked");
        }
        if (spec.protect != null) {
            if (spec.protect) {
                Protect.protect(ws, spec.password);
                done.add(spec.password != null && !spec.password.isEmpty() ? "protected with password" : "protected (no password)");
            } else {
                Protect.unprotect(ws);
                done.add("protection removed");
            }
            changed = true;
        }
        if (changed) book.pkg.markDirty(sh.part);
        return done;
    }

    // ---------------------------------------------------------------- view

    private static List<Element> sheetViews(Element ws) {
        Element views = Pkg.ensure(ws, "sheetViews", Pkg.WORKSHEET_ORDER);
        List<Element> svs = Pkg.children(views, "sheetView");
        if (svs.isEmpty()) {
            Element sv = Pkg.create(views, "sheetView");
            sv.setAttribute("workbookViewId", "0");
            views.appendChild(sv);
            svs.add(sv);
        }
        return svs;
    }

    private boolean freeze(Element ws) {
        boolean set = freezeRows > 0 || freezeCols > 0;
        if (!set && Pkg.first(ws, "sheetViews") == null) return false;
        for (Element sv : sheetViews(ws)) {
            boolean hadPane = false;
            for (Element p : Pkg.children(sv, "pane")) { sv.removeChild(p); hadPane = true; }
            if (set || hadPane) for (Element s : Pkg.children(sv, "selection")) sv.removeChild(s);
            if (!set) continue;
            String active = freezeRows > 0 && freezeCols > 0 ? "bottomRight" : freezeRows > 0 ? "bottomLeft" : "topRight";
            Element pane = Pkg.create(sv, "pane");
            if (freezeCols > 0) pane.setAttribute("xSplit", Integer.toString(freezeCols));
            if (freezeRows > 0) pane.setAttribute("ySplit", Integer.toString(freezeRows));
            pane.setAttribute("topLeftCell", freezeCell);
            pane.setAttribute("activePane", active);
            pane.setAttribute("state", "frozen");
            Element sel = Pkg.create(sv, "selection");
            sel.setAttribute("pane", active);
            sel.setAttribute("activeCell", freezeCell);
            sel.setAttribute("sqref", freezeCell);
            Pkg.insertOrdered(sv, pane, Pkg.SHEETVIEW_ORDER);
            Pkg.insertOrdered(sv, sel, Pkg.SHEETVIEW_ORDER);
        }
        return true;
    }

    private String describeFreeze() {
        if (freezeRows == 0 && freezeCols == 0) return "freeze panes removed";
        String r = freezeRows == 0 ? null : freezeRows == 1 ? "top row" : "top " + freezeRows + " rows";
        String c = freezeCols == 0 ? null : freezeCols == 1 ? "column A" : "columns A-" + Cells.colName(freezeCols);
        return "froze " + (r != null && c != null ? r + " and " + c : r != null ? r : c);
    }

    private static void gridlines(Element ws, boolean on) {
        for (Element sv : sheetViews(ws)) {
            if (on) sv.removeAttribute("showGridLines"); else sv.setAttribute("showGridLines", "0");
        }
    }

    private static void zoom(Element ws, int z) {
        for (Element sv : sheetViews(ws)) {
            sv.setAttribute("zoomScale", Integer.toString(z));
            sv.setAttribute("zoomScaleNormal", Integer.toString(z));
        }
    }

    private static void tabColor(Element ws, String argb) {
        if (argb.isEmpty()) {
            Element pr = Pkg.first(ws, "sheetPr");
            if (pr != null) for (Element t : Pkg.children(pr, "tabColor")) pr.removeChild(t);
            return;
        }
        Element pr = Pkg.ensure(ws, "sheetPr", Pkg.WORKSHEET_ORDER);
        for (Element t : Pkg.children(pr, "tabColor")) pr.removeChild(t);
        Element t = Pkg.create(pr, "tabColor");
        t.setAttribute("rgb", argb);
        Pkg.insertOrdered(pr, t, Pkg.SHEETPR_ORDER);
    }

    // ------------------------------------------------------------- columns

    /** Sets one column's width, splitting any &lt;col&gt; range that covers it. */
    private static void setWidth(Element cols, int c, double w) {
        final String ws = fmtWidth(w);
        Element sheet = (Element) cols.getParentNode();
        Grid.modifyCol(sheet, c, col -> {
            col.setAttribute("width", ws);
            col.setAttribute("customWidth", "1");
            col.removeAttribute("bestFit");
        });
    }

    private static String fmtWidth(double w) {
        BigDecimal b = new BigDecimal(w).setScale(2, RoundingMode.HALF_UP).stripTrailingZeros();
        return b.toPlainString();
    }

    /** Sizes every used column to its longest displayed value. Returns the number of columns set. */
    private int autofit(Book book, Element ws) throws XLPolishException {
        List<String> sst = book.sharedStrings();
        List<int[]> merged = mergedRects(ws);
        int[] filter = filterHeader(ws);
        TreeMap<Integer, Integer> longest = new TreeMap<>();
        for (CellRef cr : cells(ws)) {
            if (inWideMerge(merged, cr.row, cr.col)) continue;
            String text = display(cr.c, sst);
            if (text == null) continue;
            int w = Cells.displayWidth(text);
            // leave room for the filter dropdown arrow Excel draws in header cells
            if (filter != null && cr.row == filter[0] && cr.col >= filter[1] && cr.col <= filter[3]) w += FILTER_ARROW;
            Integer cur = longest.get(cr.col);
            if (cur == null || w > cur) longest.put(cr.col, w);
        }
        if (longest.isEmpty()) return 0;
        Element cols = Pkg.ensure(ws, "cols", Pkg.WORKSHEET_ORDER);
        for (Map.Entry<Integer, Integer> e : longest.entrySet()) {
            double w = Math.min(spec.autofitmax, Math.max(6.0, e.getValue() * 1.1 + 2.0));
            setWidth(cols, e.getKey(), w);
        }
        return longest.size();
    }

    /** Extra characters for the filter dropdown arrow in a header cell. */
    private static final int FILTER_ARROW = 3;

    /**
     * The filter range whose first row holds the dropdown arrows: the one this call
     * adds, or else one already in the sheet; null if the sheet will have no filter.
     */
    private int[] filterHeader(Element ws) {
        if (Boolean.FALSE.equals(spec.autofilter)) return null;
        if (Boolean.TRUE.equals(spec.autofilter)) return filterRect != null ? filterRect : usedRange(ws);
        Element af = Pkg.first(ws, "autoFilter");
        if (af == null) return null;
        try { return Cells.parseRect(af.getAttribute("ref")); } catch (XLPolishException e) { return null; }
    }

    private static List<int[]> mergedRects(Element ws) {
        List<int[]> out = new ArrayList<>();
        Element mc = Pkg.first(ws, "mergeCells");
        if (mc == null) return out;
        for (Element m : Pkg.children(mc, "mergeCell")) {
            try { out.add(Cells.parseRect(m.getAttribute("ref"))); } catch (XLPolishException ignored) { }
        }
        return out;
    }

    private static boolean inWideMerge(List<int[]> merged, int r, int c) {
        for (int[] m : merged) if (m[1] != m[3] && r >= m[0] && r <= m[2] && c >= m[1] && c <= m[3]) return true;
        return false;
    }

    /** Text a cell would show, approximately; null for empty cells. */
    private static String display(Element c, List<String> sst) {
        String t = c.getAttribute("t");
        if ("inlineStr".equals(t)) {
            Element is = Pkg.first(c, "is");
            return is == null ? null : Pkg.textOf(is);
        }
        Element v = Pkg.first(c, "v");
        if (v == null) return null;
        String s = v.getTextContent();
        if ("s".equals(t)) {
            int i = parseInt(s);
            return i >= 0 && i < sst.size() ? sst.get(i) : null;
        }
        if ("b".equals(t)) return "1".equals(s) ? "TRUE" : "FALSE";
        if ("str".equals(t) || "e".equals(t)) return s;
        try {
            double d = Double.parseDouble(s);
            if (d == Math.rint(d) && Math.abs(d) < 1e11) return Long.toString((long) d);
            if (Math.abs(d) >= 1e11 || Math.abs(d) < 1e-4) return "0.00E+00";
            String p = new BigDecimal(d).setScale(4, RoundingMode.HALF_UP).stripTrailingZeros().toPlainString();
            return p.length() > 11 ? p.substring(0, 11) : p;
        } catch (NumberFormatException e) {
            return s;
        }
    }

    // ----------------------------------------------------------- autofilter

    private String autofilter(Book book, Book.Sheet sh, Element ws, boolean on) throws XLPolishException {
        if (!on) {
            for (Element af : Pkg.children(ws, "autoFilter")) ws.removeChild(af);
            removeDefinedName(book, "_xlnm._FilterDatabase", sh.index);
            return "filter removed";
        }
        if (Pkg.first(ws, "tableParts") != null) {
            throw XLPolishException.syntax("sheet \"" + sh.name + "\" contains an Excel table, which has its own filter; "
                    + "autofilter cannot be added on top of it");
        }
        int[] rect = filterRect;
        if (rect == null) {
            rect = usedRange(ws);
            if (rect == null) throw XLPolishException.syntax("sheet \"" + sh.name + "\" has no data to filter");
            if (rect[0] == rect[2]) throw XLPolishException.syntax("sheet \"" + sh.name + "\" has only one row; nothing to filter");
        }
        for (Element af : Pkg.children(ws, "autoFilter")) ws.removeChild(af);
        Element af = Pkg.create(ws, "autoFilter");
        af.setAttribute("ref", Cells.rectRef(rect));
        Pkg.insertOrdered(ws, af, Pkg.WORKSHEET_ORDER);
        setDefinedName(book, "_xlnm._FilterDatabase", sh.index, quoteSheet(sh.name) + "!" + Cells.absRef(rect), true);
        return "filter on " + Cells.rectRef(rect);
    }

    /** Smallest rectangle holding every cell with a value; null if none. */
    private static int[] usedRange(Element ws) {
        int r1 = Integer.MAX_VALUE, c1 = Integer.MAX_VALUE, r2 = 0, c2 = 0;
        for (CellRef cr : cells(ws)) {
            if (Pkg.first(cr.c, "v") == null && Pkg.first(cr.c, "is") == null) continue;
            r1 = Math.min(r1, cr.row); c1 = Math.min(c1, cr.col);
            r2 = Math.max(r2, cr.row); c2 = Math.max(c2, cr.col);
        }
        return r2 == 0 ? null : new int[] { r1, c1, r2, c2 };
    }

    private static final class CellRef {
        final Element c; final int row, col;
        CellRef(Element c, int row, int col) { this.c = c; this.row = row; this.col = col; }
    }

    /** Every &lt;c&gt; in sheetData with its position (r attributes are optional in the format). */
    private static List<CellRef> cells(Element ws) {
        List<CellRef> out = new ArrayList<>();
        Element sd = Pkg.first(ws, "sheetData");
        if (sd == null) return out;
        int rowNo = 0;
        for (Element row : Pkg.children(sd, "row")) {
            String rr = row.getAttribute("r");
            rowNo = rr.isEmpty() ? rowNo + 1 : parseInt(rr);
            int colNo = 0;
            for (Element c : Pkg.children(row, "c")) {
                int[] p = Cells.parseCell(c.getAttribute("r"));
                colNo = p == null ? colNo + 1 : p[1];
                out.add(new CellRef(c, rowNo, colNo));
            }
        }
        return out;
    }

    // ---------------------------------------------------------------- print

    private void pageSetup(Element ws) {
        Element ps = Pkg.ensure(ws, "pageSetup", Pkg.WORKSHEET_ORDER);
        if (spec.orientation != null) ps.setAttribute("orientation", spec.orientation);
        if (spec.fitwidth) {
            ps.setAttribute("fitToWidth", "1");
            ps.setAttribute("fitToHeight", "0");
            Element pr = Pkg.ensure(ws, "sheetPr", Pkg.WORKSHEET_ORDER);
            Element pp = Pkg.ensure(pr, "pageSetUpPr", Pkg.SHEETPR_ORDER);
            pp.setAttribute("fitToPage", "1");
        }
    }

    // --------------------------------------------------------- defined names

    private static void setDefinedName(Book book, String name, int localSheetId, String value, boolean hidden)
            throws XLPolishException {
        Element root = book.wb.getDocumentElement();
        Element dns = Pkg.ensure(root, "definedNames", Pkg.WORKBOOK_ORDER);
        for (Element dn : Pkg.children(dns, "definedName")) {
            if (name.equals(dn.getAttribute("name")) && Integer.toString(localSheetId).equals(dn.getAttribute("localSheetId"))) {
                dns.removeChild(dn);
            }
        }
        Element dn = Pkg.create(dns, "definedName");
        dn.setAttribute("name", name);
        dn.setAttribute("localSheetId", Integer.toString(localSheetId));
        if (hidden) dn.setAttribute("hidden", "1");
        dn.setTextContent(value);
        dns.appendChild(dn);
        book.pkg.markDirty(book.wbPath);
    }

    private static void removeDefinedName(Book book, String name, int localSheetId) throws XLPolishException {
        Element dns = Pkg.first(book.wb.getDocumentElement(), "definedNames");
        if (dns == null) return;
        boolean any = false;
        for (Element dn : Pkg.children(dns, "definedName")) {
            if (name.equals(dn.getAttribute("name")) && Integer.toString(localSheetId).equals(dn.getAttribute("localSheetId"))) {
                dns.removeChild(dn);
                any = true;
            }
        }
        if (Pkg.children(dns, "definedName").isEmpty()) dns.getParentNode().removeChild(dns);
        if (any) book.pkg.markDirty(book.wbPath);
    }

    private static String quoteSheet(String name) { return "'" + name.replace("'", "''") + "'"; }

    // -------------------------------------------------- conditional formats

    private static int maxPriority(Element ws) {
        int max = 0;
        for (Element cf : Pkg.children(ws, "conditionalFormatting")) {
            for (Element r : Pkg.children(cf, "cfRule")) max = Math.max(max, parseInt(r.getAttribute("priority")));
        }
        return max;
    }

    private static Element newRule(Element ws, String sqref, String type, int priority) {
        Element cf = Pkg.create(ws, "conditionalFormatting");
        cf.setAttribute("sqref", sqref);
        Element rule = Pkg.create(cf, "cfRule");
        rule.setAttribute("type", type);
        rule.setAttribute("priority", Integer.toString(priority));
        cf.appendChild(rule);
        Pkg.insertOrdered(ws, cf, Pkg.WORKSHEET_ORDER);
        return rule;
    }

    private static Element cfvo(Element parent, String type, String val) {
        Element v = Pkg.create(parent, "cfvo");
        v.setAttribute("type", type);
        if (val != null) v.setAttribute("val", val);
        parent.appendChild(v);
        return v;
    }

    private static void color(Element parent, String argb) {
        Element c = Pkg.create(parent, "color");
        c.setAttribute("rgb", argb);
        parent.appendChild(c);
    }

    private static void addColorScale(Element ws, ColorScale cs, int priority) {
        Element rule = newRule(ws, cs.sqref, "colorScale", priority);
        Element sc = Pkg.create(rule, "colorScale");
        rule.appendChild(sc);
        cfvo(sc, "min", null);
        if (cs.colors.length == 3) cfvo(sc, "percentile", "50");
        cfvo(sc, "max", null);
        for (String c : cs.colors) color(sc, c);
    }

    private static void addDataBar(Element ws, DataBar d, int priority) {
        Element rule = newRule(ws, d.sqref, "dataBar", priority);
        Element db = Pkg.create(rule, "dataBar");
        rule.appendChild(db);
        cfvo(db, "min", null);
        cfvo(db, "max", null);
        color(db, d.color);
    }

    private static void addHighlight(Book book, Element ws, Highlight h, int priority) throws XLPolishException {
        int dxf = addDxf(book, h.fill, h.font);
        // A formula rule rather than Excel's "cell value" rule, because the latter
        // treats empty cells as 0 (so "<= 1" would highlight every missing value).
        // The formula is written for the top-left cell; Excel shifts it for the rest.
        Element rule = newRule(ws, h.sqref, "expression", priority);
        rule.setAttribute("dxfId", Integer.toString(dxf));
        Element f = Pkg.create(rule, "formula");
        f.setTextContent("AND(NOT(ISBLANK(" + h.anchor + "))," + h.anchor + h.op + h.formula + ")");
        rule.appendChild(f);
    }

    /** Adds a differential format (fill and optional font colour) to styles.xml; returns its index. */
    private static int addDxf(Book book, String fill, String font) throws XLPolishException {
        if (book.stylesPath == null || !book.pkg.has(book.stylesPath)) {
            throw new XLPolishException(610, "workbook has no styles part; open and re-save it in Excel, then try again");
        }
        Element root = book.pkg.doc(book.stylesPath).getDocumentElement();
        Element dxfs = Pkg.ensure(root, "dxfs", Pkg.STYLESHEET_ORDER);
        Element dxf = Pkg.create(dxfs, "dxf");
        if (font != null) {
            Element f = Pkg.create(dxf, "font");
            color(f, font);
            dxf.appendChild(f);
        }
        Element fl = Pkg.create(dxf, "fill");
        Element pf = Pkg.create(fl, "patternFill");
        Element bg = Pkg.create(pf, "bgColor");
        bg.setAttribute("rgb", fill);
        pf.appendChild(bg);
        fl.appendChild(pf);
        dxf.appendChild(fl);
        dxfs.appendChild(dxf);
        int n = Pkg.children(dxfs, "dxf").size();
        dxfs.setAttribute("count", Integer.toString(n));
        book.pkg.markDirty(book.stylesPath);
        return n - 1;
    }

    // ------------------------------------------------------------- dropdowns

    private static void addDropdown(Element ws, Dropdown d) {
        Element dvs = Pkg.ensure(ws, "dataValidations", Pkg.WORKSHEET_ORDER);
        Element dv = Pkg.create(dvs, "dataValidation");
        dv.setAttribute("type", "list");
        dv.setAttribute("allowBlank", "1");
        dv.setAttribute("showInputMessage", "1");
        dv.setAttribute("showErrorMessage", "1");
        dv.setAttribute("sqref", d.sqref);
        Element f = Pkg.create(dv, "formula1");
        f.setTextContent(d.formula);
        dv.appendChild(f);
        dvs.appendChild(dv);
        dvs.setAttribute("count", Integer.toString(Pkg.children(dvs, "dataValidation").size()));
    }

    // ------------------------------------------------------- sheet visibility

    private static List<Element> sheetElements(Book book) {
        return Pkg.descendants(book.wb.getDocumentElement(), "sheet");
    }

    private static boolean visible(Element sheetEl) {
        String st = sheetEl.getAttribute("state");
        return st.isEmpty() || "visible".equals(st);
    }

    private static String setVisible(Book book, Book.Sheet sh, boolean show) throws XLPolishException {
        Element el = sheetElements(book).get(sh.index);
        if (show) {
            el.removeAttribute("state");
        } else {
            int visibleOthers = 0;
            List<Element> all = sheetElements(book);
            for (int i = 0; i < all.size(); i++) if (i != sh.index && visible(all.get(i))) visibleOthers++;
            if (visibleOthers == 0) {
                throw XLPolishException.syntax("cannot hide \"" + sh.name + "\": a workbook needs at least one visible sheet");
            }
            el.setAttribute("state", "hidden");
        }
        book.pkg.markDirty(book.wbPath);
        return show ? "sheet shown" : "sheet hidden";
    }

    /** If the active or first-shown tab is now hidden, move both to the first visible sheet. */
    private static void fixActiveTab(Book book) throws XLPolishException {
        List<Element> all = sheetElements(book);
        int firstVisible = -1;
        for (int i = 0; i < all.size(); i++) if (visible(all.get(i))) { firstVisible = i; break; }
        Element views = Pkg.first(book.wb.getDocumentElement(), "bookViews");
        if (views == null || firstVisible < 0) return;
        for (Element wv : Pkg.children(views, "workbookView")) {
            int active = wv.hasAttribute("activeTab") ? Grid.toInt(wv.getAttribute("activeTab")) : 0;
            if (active < 0 || active >= all.size() || !visible(all.get(active))) wv.setAttribute("activeTab", Integer.toString(firstVisible));
            int first = wv.hasAttribute("firstSheet") ? Grid.toInt(wv.getAttribute("firstSheet")) : 0;
            if (first < 0 || first >= all.size() || !visible(all.get(first))) wv.setAttribute("firstSheet", Integer.toString(firstVisible));
        }
        book.pkg.markDirty(book.wbPath);
        // a hidden sheet must not be the "selected" tab; select the new active one instead
        int active = Grid.toInt(Pkg.children(views, "workbookView").isEmpty() ? "0"
                : Pkg.children(views, "workbookView").get(0).getAttribute("activeTab"));
        for (Book.Sheet s : book.sheets) {
            if (!s.worksheet || !book.pkg.has(s.part)) continue;
            Element ws = book.pkg.doc(s.part).getDocumentElement();
            Element sv = Pkg.first(ws, "sheetViews");
            if (sv == null) continue;
            for (Element v : Pkg.children(sv, "sheetView")) {
                boolean want = s.index == active;
                boolean has = "1".equals(v.getAttribute("tabSelected")) || "true".equals(v.getAttribute("tabSelected"));
                if (want && !has) { v.setAttribute("tabSelected", "1"); book.pkg.markDirty(s.part); }
                if (!want && has) { v.removeAttribute("tabSelected"); book.pkg.markDirty(s.part); }
            }
        }
    }

    // ---------------------------------------------------------------- misc

    private static int parseInt(String s) {
        try { return Integer.parseInt(s.trim()); } catch (Exception e) { return -1; }
    }
}
