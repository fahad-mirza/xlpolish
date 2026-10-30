package xlpolish;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;
import java.util.function.IntFunction;
import java.util.function.IntUnaryOperator;

import org.w3c.dom.Element;

/**
 * Row and column grouping (Excel's Data &gt; Group). Each group raises the outline
 * level of its rows or columns by one, so overlapping groups nest, up to Excel's
 * limit of 7 levels. With collapse, the grouped rows or columns start hidden and
 * the row below (or column right of) each group carries the "collapsed" mark that
 * shows the + button.
 */
final class Outline {

    static final int MAX_LEVEL = 7;

    private Outline() { }

    /**
     * Sets Excel's "Summary rows below detail" / "Summary columns to right of detail"
     * checkboxes (outlinePr). above/left = null leaves that setting alone.
     */
    static void summarySide(Element ws, Boolean above, Boolean left) {
        Element pr = Pkg.ensure(ws, "sheetPr", Pkg.WORKSHEET_ORDER);
        Element op = Pkg.ensure(pr, "outlinePr", Pkg.SHEETPR_ORDER);
        if (above != null) { if (above) op.setAttribute("summaryBelow", "0"); else op.removeAttribute("summaryBelow"); }
        if (left != null) { if (left) op.setAttribute("summaryRight", "0"); else op.removeAttribute("summaryRight"); }
        if (!op.hasAttributes() && !op.hasChildNodes()) pr.removeChild(op);
    }

    /** True if this sheet puts summary rows above detail (button before each row group). */
    static boolean summaryAbove(Element ws) {
        Element pr = Pkg.first(ws, "sheetPr");
        Element op = pr == null ? null : Pkg.first(pr, "outlinePr");
        String v = op == null ? "" : op.getAttribute("summaryBelow");
        return "0".equals(v) || "false".equals(v);
    }

    /** True if this sheet puts summary columns left of detail (button before each column group). */
    static boolean summaryLeft(Element ws) {
        Element pr = Pkg.first(ws, "sheetPr");
        Element op = pr == null ? null : Pkg.first(pr, "outlinePr");
        String v = op == null ? "" : op.getAttribute("summaryRight");
        return "0".equals(v) || "false".equals(v);
    }

    static List<String> groupRows(Element ws, List<int[]> ranges, boolean collapse) throws XLPolishException {
        Grid g = new Grid(ws);
        boolean above = summaryAbove(ws);
        for (int[] rg : ranges) {
            for (int r = rg[0]; r <= rg[1]; r++) {
                Element row = g.row(r);
                int lvl = Grid.toInt(row.getAttribute("outlineLevel")) + 1;
                if (lvl > MAX_LEVEL) throw XLPolishException.syntax("row " + r + " would be nested more than "
                        + MAX_LEVEL + " groups deep, which is Excel's limit");
                row.setAttribute("outlineLevel", Integer.toString(lvl));
                if (collapse) row.setAttribute("hidden", "1");
            }
            // the +/- button (and Excel's "collapsed" mark) sits on the summary row
            if (collapse && !above && rg[1] < Cells.MAX_ROWS) g.row(rg[1] + 1).setAttribute("collapsed", "1");
            if (collapse && above && rg[0] > 1) g.row(rg[0] - 1).setAttribute("collapsed", "1");
        }
        int max = 0;
        for (Element row : g.rows().values()) max = Math.max(max, Grid.toInt(row.getAttribute("outlineLevel")));
        Element fmt = formatPr(ws);
        if (max > 0) fmt.setAttribute("outlineLevelRow", Integer.toString(max)); else fmt.removeAttribute("outlineLevelRow");

        final Map<Integer, Element> rows = g.rows();
        return touching(ranges, Cells.MAX_ROWS, i -> {
            Element row = rows.get(i);
            return row == null ? 0 : Grid.toInt(row.getAttribute("outlineLevel"));
        }, "rows", Integer::toString);
    }

    static List<String> groupCols(Element ws, List<int[]> ranges, final boolean collapse) throws XLPolishException {
        boolean left = summaryLeft(ws);
        for (int[] rg : ranges) {
            for (int c = rg[0]; c <= rg[1]; c++) {
                final int cc = c;
                final XLPolishException[] err = new XLPolishException[1];
                Grid.modifyCol(ws, c, col -> {
                    int lvl = Grid.toInt(col.getAttribute("outlineLevel")) + 1;
                    if (lvl > MAX_LEVEL) {
                        err[0] = XLPolishException.syntax("column " + Cells.colName(cc) + " would be nested more than "
                                + MAX_LEVEL + " groups deep, which is Excel's limit");
                        return;
                    }
                    col.setAttribute("outlineLevel", Integer.toString(lvl));
                    if (collapse) col.setAttribute("hidden", "1");
                });
                if (err[0] != null) throw err[0];
            }
            if (collapse && !left && rg[1] < Cells.MAX_COLS) Grid.modifyCol(ws, rg[1] + 1, col -> col.setAttribute("collapsed", "1"));
            if (collapse && left && rg[0] > 1) Grid.modifyCol(ws, rg[0] - 1, col -> col.setAttribute("collapsed", "1"));
        }
        int max = 0;
        for (Element col : Grid.cols(ws)) max = Math.max(max, Grid.toInt(col.getAttribute("outlineLevel")));
        Element fmt = formatPr(ws);
        if (max > 0) fmt.setAttribute("outlineLevelCol", Integer.toString(max)); else fmt.removeAttribute("outlineLevelCol");

        final Map<Integer, Integer> levels = new HashMap<>();
        for (Element col : Grid.cols(ws)) {
            int lvl = Grid.toInt(col.getAttribute("outlineLevel"));
            for (int c = Grid.toInt(col.getAttribute("min")); c <= Grid.toInt(col.getAttribute("max")); c++) levels.put(c, lvl);
        }
        return touching(ranges, Cells.MAX_COLS, i -> levels.containsKey(i) ? levels.get(i) : 0, "columns", Cells::colName);
    }

    /**
     * Excel stores grouping only as a level number per row or column, so two groups
     * at the same level with no ungrouped row/column between them are one group to
     * Excel (and the first group's -/+ button would fall inside the second). Returns
     * one plain-language warning per place where a new group touches another.
     */
    private static List<String> touching(List<int[]> ranges, int limit, IntUnaryOperator level, String what,
                                         IntFunction<String> name) {
        TreeSet<Integer> seams = new TreeSet<>();      // index i: the boundary between i and i + 1
        for (int[] rg : ranges) {
            if (rg[0] > 1 && level.applyAsInt(rg[0] - 1) >= level.applyAsInt(rg[0])) seams.add(rg[0] - 1);
            if (rg[1] < limit && level.applyAsInt(rg[1] + 1) >= level.applyAsInt(rg[1])) seams.add(rg[1]);
        }
        List<String> out = new ArrayList<>();
        for (int i : seams) {
            out.add("note: grouped " + what + " " + name.apply(i) + " and " + name.apply(i + 1)
                    + " touch, so Excel shows the groups on either side as one group; leave at least one ungrouped "
                    + (what.equals("rows") ? "row" : "column") + " between groups to keep them separate");
        }
        return out;
    }

    /** Removes all grouping; rows and columns hidden by a collapsed group are shown again. */
    static boolean ungroup(Element ws) {
        boolean any = false;
        Element sd = Pkg.first(ws, "sheetData");
        if (sd != null) {
            for (Element row : Pkg.children(sd, "row")) {
                if (row.hasAttribute("outlineLevel")) {
                    row.removeAttribute("outlineLevel");
                    row.removeAttribute("hidden");
                    any = true;
                }
                if (row.hasAttribute("collapsed")) { row.removeAttribute("collapsed"); any = true; }
            }
        }
        for (Element col : Grid.cols(ws)) {
            if (col.hasAttribute("outlineLevel")) {
                col.removeAttribute("outlineLevel");
                col.removeAttribute("hidden");
                any = true;
            }
            if (col.hasAttribute("collapsed")) { col.removeAttribute("collapsed"); any = true; }
        }
        Element fmt = Pkg.first(ws, "sheetFormatPr");
        if (fmt != null) { fmt.removeAttribute("outlineLevelRow"); fmt.removeAttribute("outlineLevelCol"); }
        return any;
    }

    /** &lt;sheetFormatPr&gt;, created with Excel's default row height if missing (that attribute is required). */
    private static Element formatPr(Element ws) {
        Element fmt = Pkg.first(ws, "sheetFormatPr");
        if (fmt == null) {
            fmt = Pkg.create(ws, "sheetFormatPr");
            fmt.setAttribute("defaultRowHeight", "15");
            Pkg.insertOrdered(ws, fmt, Pkg.WORKSHEET_ORDER);
        }
        return fmt;
    }
}
