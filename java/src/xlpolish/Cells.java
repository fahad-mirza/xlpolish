package xlpolish;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Cell-reference helpers: A1 notation, columns, ranges. */
final class Cells {
    static final int MAX_ROWS = 1048576;
    static final int MAX_COLS = 16384;

    private static final Pattern CELL = Pattern.compile("^\\$?([A-Z]{1,3})\\$?([0-9]+)$");
    private static final Pattern COLS = Pattern.compile("^\\$?([A-Z]{1,3}):\\$?([A-Z]{1,3})$");
    private static final Pattern ROWS = Pattern.compile("^\\$?([0-9]+):\\$?([0-9]+)$");

    private Cells() { }

    static int colNum(String letters) {
        int c = 0;
        for (int i = 0; i < letters.length(); i++) c = c * 26 + (letters.charAt(i) - 'A' + 1);
        return c;
    }

    static String colName(int n) {
        StringBuilder sb = new StringBuilder();
        while (n > 0) { int m = (n - 1) % 26; sb.insert(0, (char) ('A' + m)); n = (n - 1) / 26; }
        return sb.toString();
    }

    static String ref(int row, int col) { return colName(col) + row; }

    /** "B2" -> {row 2, col 2}; null if not a single cell. */
    static int[] parseCell(String s) {
        if (s == null) return null;
        Matcher m = CELL.matcher(s.trim().toUpperCase(Locale.ROOT));
        if (!m.matches()) return null;
        int col = colNum(m.group(1));
        long row;
        try { row = Long.parseLong(m.group(2)); } catch (NumberFormatException e) { return null; }
        if (col < 1 || col > MAX_COLS || row < 1 || row > MAX_ROWS) return null;
        return new int[] { (int) row, col };
    }

    /** One rectangle {r1, c1, r2, c2} from "A1", "A1:D9", "B:D" or "2:5". */
    static int[] parseRect(String s) throws XLPolishException {
        String u = s.trim().toUpperCase(Locale.ROOT);
        Matcher m;
        if ((m = COLS.matcher(u)).matches()) {
            int a = colNum(m.group(1)), b = colNum(m.group(2));
            if (a < 1 || b > MAX_COLS || a > b) throw XLPolishException.syntax("invalid column range " + s);
            return new int[] { 1, a, MAX_ROWS, b };
        }
        if ((m = ROWS.matcher(u)).matches()) {
            long a = Long.parseLong(m.group(1)), b = Long.parseLong(m.group(2));
            if (a < 1 || b > MAX_ROWS || a > b) throw XLPolishException.syntax("invalid row range " + s);
            return new int[] { (int) a, 1, (int) b, MAX_COLS };
        }
        String[] p = u.split(":", -1);
        if (p.length == 1 || p.length == 2) {
            int[] x = parseCell(p[0]);
            int[] y = p.length == 2 ? parseCell(p[1]) : x;
            if (x != null && y != null) {
                return new int[] { Math.min(x[0], y[0]), Math.min(x[1], y[1]), Math.max(x[0], y[0]), Math.max(x[1], y[1]) };
            }
        }
        throw XLPolishException.syntax("\"" + s + "\" is not a valid cell range (examples: B2:H164, B:D, 2:5)");
    }

    static String rectRef(int[] r) {
        String a = ref(r[0], r[1]), b = ref(r[2], r[3]);
        return a.equals(b) ? a : a + ":" + b;
    }

    /** Absolute form for defined names: $A$1:$H$164. */
    static String absRef(int[] r) {
        return "$" + colName(r[1]) + "$" + r[0] + ":$" + colName(r[3]) + "$" + r[2];
    }

    /** A space-separated list of ranges, normalised for an sqref attribute. */
    static String sqref(String s) throws XLPolishException {
        List<String> out = new ArrayList<>();
        for (String piece : s.trim().split("[\\s,]+")) {
            if (!piece.isEmpty()) out.add(rectRef(parseRect(piece)));
        }
        if (out.isEmpty()) throw XLPolishException.syntax("a cell range is required");
        return String.join(" ", out);
    }

    /** Display width of a string in "character" units; East Asian wide characters count as 2. */
    static int displayWidth(String s) {
        int w = 0, best = 0;
        for (int i = 0; i < s.length(); ) {
            int cp = s.codePointAt(i);
            i += Character.charCount(cp);
            if (cp == '\n') { best = Math.max(best, w); w = 0; continue; }
            w += isWide(cp) ? 2 : 1;
        }
        return Math.max(best, w);
    }

    private static boolean isWide(int cp) {
        return (cp >= 0x1100 && cp <= 0x115F) || (cp >= 0x2E80 && cp <= 0xA4CF) || (cp >= 0xAC00 && cp <= 0xD7A3)
                || (cp >= 0xF900 && cp <= 0xFAFF) || (cp >= 0xFE30 && cp <= 0xFE4F) || (cp >= 0xFF00 && cp <= 0xFF60)
                || (cp >= 0xFFE0 && cp <= 0xFFE6) || (cp >= 0x20000 && cp <= 0x3FFFD);
    }
}
