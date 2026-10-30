package xlpolish;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.w3c.dom.Element;

/**
 * Sheet protection and unlocked cells.
 *
 * Protection: users can still select cells and use filter dropdowns and sorting;
 * everything else (editing locked cells, inserting rows, formatting) is blocked.
 * A password is stored the way Excel 2013+ stores it: SHA-512 over a random
 * 16-byte salt and the UTF-16LE password, re-hashed 100,000 times with the
 * iteration number appended, then Base64-encoded. The password itself is never
 * written to the file.
 *
 * Unlocking: every cell is "locked" by default, which only matters once the sheet
 * is protected. unlock() gives the chosen cells a copy of their current style with
 * the locked flag off, so they stay editable on a protected sheet.
 */
final class Protect {

    static final int SPIN = 100000;
    private static final long MAX_CELLS = 2_000_000L;

    private Protect() { }

    static void protect(Element ws, String password) throws XLPolishException {
        unprotect(ws);
        Element sp = Pkg.create(ws, "sheetProtection");
        if (password != null && !password.isEmpty()) {
            byte[] salt = new byte[16];
            new SecureRandom().nextBytes(salt);
            sp.setAttribute("algorithmName", "SHA-512");
            sp.setAttribute("hashValue", Base64.getEncoder().encodeToString(hash(password, salt, SPIN)));
            sp.setAttribute("saltValue", Base64.getEncoder().encodeToString(salt));
            sp.setAttribute("spinCount", Integer.toString(SPIN));
        }
        sp.setAttribute("sheet", "1");
        sp.setAttribute("objects", "1");
        sp.setAttribute("scenarios", "1");
        sp.setAttribute("autoFilter", "0");   // 0 = allowed while protected
        sp.setAttribute("sort", "0");
        Pkg.insertOrdered(ws, sp, Pkg.WORKSHEET_ORDER);
    }

    static boolean unprotect(Element ws) {
        boolean any = false;
        for (Element sp : Pkg.children(ws, "sheetProtection")) { ws.removeChild(sp); any = true; }
        return any;
    }

    /** SHA-512 password hash as used by Excel sheet protection. */
    static byte[] hash(String password, byte[] salt, int spin) throws XLPolishException {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-512");
            md.update(salt);
            md.update(password.getBytes(StandardCharsets.UTF_16LE));
            byte[] h = md.digest();
            byte[] it = new byte[4];
            for (int i = 0; i < spin; i++) {
                it[0] = (byte) i; it[1] = (byte) (i >>> 8); it[2] = (byte) (i >>> 16); it[3] = (byte) (i >>> 24);
                md.update(h);
                md.update(it);
                h = md.digest();
            }
            return h;
        } catch (Exception e) {
            throw new XLPolishException(198, "could not hash the password: " + e.getMessage());
        }
    }

    // ------------------------------------------------------------- unlocking

    /** Makes the cells in rects editable on a protected sheet. */
    static void unlock(Book book, Element ws, List<int[]> rects) throws XLPolishException {
        if (book.stylesPath == null || !book.pkg.has(book.stylesPath)) {
            throw new XLPolishException(610, "workbook has no styles part; open and re-save it in Excel, then try again");
        }
        final Element styles = book.pkg.doc(book.stylesPath).getDocumentElement();
        final Element xfs = Pkg.first(styles, "cellXfs");
        if (xfs == null || Pkg.children(xfs, "xf").isEmpty()) {
            throw new XLPolishException(610, "workbook has no cell styles; open and re-save it in Excel, then try again");
        }
        final Map<Integer, Integer> map = new HashMap<>();
        Grid g = new Grid(ws);

        for (int[] r : rects) {
            boolean allRows = r[0] == 1 && r[2] == Cells.MAX_ROWS;
            boolean allCols = r[1] == 1 && r[3] == Cells.MAX_COLS;
            if (allRows) {
                // whole columns: the column style covers cells typed in later
                for (int c = r[1]; c <= r[3]; c++) {
                    Grid.modifyCol(ws, c, col -> col.setAttribute("style",
                            Integer.toString(unlockedXf(xfs, map, Grid.toInt(col.getAttribute("style"))))));
                }
                for (Element row : g.rows().values()) {
                    for (Map.Entry<Integer, Element> e : Grid.cellsOf(row).entrySet()) {
                        if (e.getKey() >= r[1] && e.getKey() <= r[3]) restyle(xfs, map, e.getValue());
                    }
                }
            } else if (allCols) {
                // whole rows: the row style covers cells typed in later
                for (int rr = r[0]; rr <= r[2]; rr++) {
                    Element row = g.row(rr);
                    row.setAttribute("s", Integer.toString(unlockedXf(xfs, map, Grid.toInt(row.getAttribute("s")))));
                    row.setAttribute("customFormat", "1");
                    for (Element c : Grid.cellsOf(row).values()) restyle(xfs, map, c);
                }
            } else {
                long n = (long) (r[2] - r[0] + 1) * (r[3] - r[1] + 1);
                if (n > MAX_CELLS) {
                    throw XLPolishException.syntax("unlock() range " + Cells.rectRef(r) + " has " + n + " cells; "
                            + "for large areas give whole columns (e.g. D:F) or whole rows (e.g. 2:5000) instead");
                }
                for (int rr = r[0]; rr <= r[2]; rr++) {
                    for (int c = r[1]; c <= r[3]; c++) restyle(xfs, map, g.cell(rr, c));
                }
            }
        }
        xfs.setAttribute("count", Integer.toString(Pkg.children(xfs, "xf").size()));
        book.pkg.markDirty(book.stylesPath);
    }

    private static void restyle(Element xfs, Map<Integer, Integer> map, Element cell) {
        int s = cell.hasAttribute("s") ? Grid.toInt(cell.getAttribute("s")) : 0;
        cell.setAttribute("s", Integer.toString(unlockedXf(xfs, map, s)));
    }

    /** Index of a copy of cell style s with the locked flag off (created once per style). */
    private static int unlockedXf(Element xfs, Map<Integer, Integer> map, int s) {
        Integer done = map.get(s);
        if (done != null) return done;
        List<Element> all = Pkg.children(xfs, "xf");
        Element src = all.get(s >= 0 && s < all.size() ? s : 0);
        // already unlocked? reuse it
        Element p0 = Pkg.first(src, "protection");
        if (p0 != null && "0".equals(p0.getAttribute("locked"))) { map.put(s, s); return s; }
        Element copy = (Element) src.cloneNode(true);
        for (Element p : Pkg.children(copy, "protection")) copy.removeChild(p);
        Element prot = Pkg.create(copy, "protection");
        prot.setAttribute("locked", "0");
        Pkg.insertOrdered(copy, prot, Pkg.XF_ORDER);
        copy.setAttribute("applyProtection", "1");
        xfs.appendChild(copy);
        int idx = all.size();
        map.put(s, idx);
        map.put(idx, idx);
        return idx;
    }

    /** For tests: the hash Excel would store, given a salt. */
    static String hashBase64(String password, byte[] salt) throws XLPolishException {
        return Base64.getEncoder().encodeToString(hash(password, salt, SPIN));
    }
}
