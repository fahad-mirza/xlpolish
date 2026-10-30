package xlpolish;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.function.Consumer;

import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * Finds or creates rows and cells in a worksheet's &lt;sheetData&gt;, and column
 * entries in &lt;cols&gt;, keeping them in the ascending order Excel requires.
 */
final class Grid {

    final Element ws;
    final Element sheetData;
    private final TreeMap<Integer, Element> rows = new TreeMap<>();

    Grid(Element ws) {
        this.ws = ws;
        this.sheetData = Pkg.ensure(ws, "sheetData", Pkg.WORKSHEET_ORDER);
        int rowNo = 0;
        for (Element row : Pkg.children(sheetData, "row")) {
            String rr = row.getAttribute("r");
            rowNo = rr.isEmpty() ? rowNo + 1 : Integer.parseInt(rr.trim());
            if (rr.isEmpty()) row.setAttribute("r", Integer.toString(rowNo));   // r is optional; make it explicit
            rows.put(rowNo, row);
        }
    }

    /** All existing rows, in order. */
    Map<Integer, Element> rows() { return rows; }

    /** The row element for row r, created (empty) in order if missing. */
    Element row(int r) {
        Element row = rows.get(r);
        if (row != null) return row;
        row = Pkg.create(sheetData, "row");
        row.setAttribute("r", Integer.toString(r));
        Map.Entry<Integer, Element> next = rows.higherEntry(r);
        sheetData.insertBefore(row, next == null ? null : next.getValue());
        rows.put(r, row);
        return row;
    }

    /** Existing cells of a row as (column -> element), assigning r attributes where missing. */
    static TreeMap<Integer, Element> cellsOf(Element row) {
        TreeMap<Integer, Element> out = new TreeMap<>();
        int r = Integer.parseInt(row.getAttribute("r").trim());
        int colNo = 0;
        for (Element c : Pkg.children(row, "c")) {
            int[] p = Cells.parseCell(c.getAttribute("r"));
            colNo = p == null ? colNo + 1 : p[1];
            if (p == null) c.setAttribute("r", Cells.ref(r, colNo));
            out.put(colNo, c);
        }
        return out;
    }

    /** The cell element at (r, c), created (empty) in order if missing. */
    Element cell(int r, int c) {
        Element row = row(r);
        TreeMap<Integer, Element> cs = cellsOf(row);
        Element cell = cs.get(c);
        if (cell != null) return cell;
        cell = Pkg.create(row, "c");
        cell.setAttribute("r", Cells.ref(r, c));
        Map.Entry<Integer, Element> next = cs.higherEntry(c);
        row.insertBefore(cell, next == null ? null : next.getValue());
        // keep the row's optional "spans" hint from contradicting its cells
        row.removeAttribute("spans");
        return cell;
    }

    // ---------------------------------------------------------------- columns

    /** Default column width Excel writes for its standard 8.43-character column. */
    static final String DEFAULT_COL_WIDTH = "9.140625";

    /**
     * Applies f to the &lt;col&gt; entry for column c alone, first splitting any
     * entry that covers a wider range, or creating one if none covers c.
     */
    static void modifyCol(Element ws, int c, Consumer<Element> f) {
        Element cols = Pkg.ensure(ws, "cols", Pkg.WORKSHEET_ORDER);
        for (Element col : Pkg.children(cols, "col")) {
            int min = toInt(col.getAttribute("min")), max = toInt(col.getAttribute("max"));
            if (c < min || c > max) continue;
            if (min < c) {
                Element left = (Element) col.cloneNode(true);
                left.setAttribute("max", Integer.toString(c - 1));
                cols.insertBefore(left, col);
            }
            if (c < max) {
                Element right = (Element) col.cloneNode(true);
                right.setAttribute("min", Integer.toString(c + 1));
                cols.insertBefore(right, col.getNextSibling());
            }
            col.setAttribute("min", Integer.toString(c));
            col.setAttribute("max", Integer.toString(c));
            f.accept(col);
            return;
        }
        Element nc = Pkg.create(cols, "col");
        nc.setAttribute("min", Integer.toString(c));
        nc.setAttribute("max", Integer.toString(c));
        nc.setAttribute("width", DEFAULT_COL_WIDTH);
        Node before = null;
        for (Element col : Pkg.children(cols, "col")) {
            if (toInt(col.getAttribute("min")) > c) { before = col; break; }
        }
        cols.insertBefore(nc, before);
        f.accept(nc);
    }

    /** Column entries currently in the sheet. */
    static List<Element> cols(Element ws) {
        Element cols = Pkg.first(ws, "cols");
        return cols == null ? new ArrayList<Element>() : Pkg.children(cols, "col");
    }

    static int toInt(String s) {
        try { return Integer.parseInt(s.trim()); } catch (Exception e) { return 0; }
    }
}
