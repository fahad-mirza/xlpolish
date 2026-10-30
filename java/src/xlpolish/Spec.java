package xlpolish;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Options for one xlpolish call. The ado-file writes them to a small UTF-8 text
 * file, one "key&lt;TAB&gt;value" per line, so values may contain spaces and quotes.
 */
final class Spec {

    enum Mode { FIRST, NAMED, ALL }

    String file;
    Mode mode = Mode.FIRST;
    String sheet;

    String freeze;                 // cell; "A1" removes
    Boolean gridlines;             // null = leave
    Integer zoom;
    String tabcolor;               // colour spec or "none"
    Boolean autofilter;            // null leave, TRUE on, FALSE remove
    String filterrange;
    boolean autofit;
    double autofitmax = 60;
    String colwidth;               // "A=30 B:D=12"
    String orientation;            // landscape | portrait
    boolean fitwidth;
    String printtitles;            // "1" / "1:2" / "A:A" / "1:2 A:A"
    final List<String> colorscale = new ArrayList<>();
    final List<String> databar = new ArrayList<>();
    final List<String> highlight = new ArrayList<>();
    final List<String> dropdown = new ArrayList<>();
    final List<String> comment = new ArrayList<>();
    String commentAuthor = "";
    boolean removeComments;
    String groupRows, groupCols;
    boolean collapse, ungroup;
    Boolean protect;               // null leave, TRUE protect, FALSE unprotect
    String password;
    final List<String> unlock = new ArrayList<>();
    Boolean hideSheet;             // null leave, TRUE hide, FALSE show
    Boolean summaryAbove;          // null leave, TRUE above detail, FALSE below (Excel default)
    Boolean summaryLeft;           // null leave, TRUE left of detail, FALSE right (Excel default)

    static Spec read(Path p) throws XLPolishException {
        List<String> lines;
        try {
            lines = Files.readAllLines(p, StandardCharsets.UTF_8);
        } catch (IOException e) {
            throw new XLPolishException(603, "could not read option file " + p + ": " + e.getMessage());
        }
        Spec s = new Spec();
        for (String line : lines) {
            if (line.isEmpty()) continue;
            if (line.charAt(0) == '\uFEFF') line = line.substring(1);
            int tab = line.indexOf('\t');
            String k = (tab < 0 ? line : line.substring(0, tab)).trim();
            String v = tab < 0 ? "" : line.substring(tab + 1).trim();
            s.set(k, v);
        }
        if (s.file == null || s.file.isEmpty()) throw XLPolishException.syntax("no file given");
        return s;
    }

    void set(String k, String v) throws XLPolishException {
        switch (k) {
            case "file": file = v; break;
            case "sheet": sheet = v; mode = Mode.NAMED; break;
            case "allsheets": mode = Mode.ALL; break;
            case "freeze": freeze = v; break;
            case "gridlines": gridlines = parseOnOff(k, v); break;
            case "zoom":
                try { zoom = Integer.parseInt(v); } catch (NumberFormatException e) { zoom = -1; }
                if (zoom < 10 || zoom > 400) throw XLPolishException.syntax("zoom() must be a whole number from 10 to 400");
                break;
            case "tabcolor": tabcolor = v; break;
            case "autofilter": autofilter = parseOnOff(k, v); break;
            case "filterrange": filterrange = v; autofilter = Boolean.TRUE; break;
            case "autofit": autofit = true; break;
            case "autofitmax":
                try { autofitmax = Double.parseDouble(v); } catch (NumberFormatException e) { autofitmax = -1; }
                if (autofitmax < 1 || autofitmax > 255) throw XLPolishException.syntax("autofitmax() must be between 1 and 255");
                break;
            case "colwidth": colwidth = v; break;
            case "orientation": orientation = v.toLowerCase(Locale.ROOT); break;
            case "fitwidth": fitwidth = true; break;
            case "printtitles": printtitles = v; break;
            case "colorscale": addRules(colorscale, v); break;
            case "databar": addRules(databar, v); break;
            case "highlight": addRules(highlight, v); break;
            case "dropdown": addRules(dropdown, v); break;
            case "comment": addRules(comment, v); break;
            case "commentauthor": commentAuthor = v; break;
            case "removecomments": removeComments = true; break;
            case "grouprows": groupRows = v; break;
            case "groupcols": groupCols = v; break;
            case "collapse": collapse = true; break;
            case "ungroup": ungroup = true; break;
            case "protect": protect = Boolean.TRUE; break;
            case "unprotect": protect = Boolean.FALSE; break;
            case "password": password = v; break;
            case "unlock": addRules(unlock, v); break;
            case "hidesheet": hideSheet = Boolean.TRUE; break;
            case "showsheet": hideSheet = Boolean.FALSE; break;
            case "summaryrows": summaryAbove = "above".equals(v); break;
            case "summarycols": summaryLeft = "left".equals(v); break;
            default: throw XLPolishException.syntax("unknown option " + k);
        }
    }

    private static Boolean parseOnOff(String k, String v) throws XLPolishException {
        if ("on".equals(v)) return Boolean.TRUE;
        if ("off".equals(v)) return Boolean.FALSE;
        throw XLPolishException.syntax("internal: " + k + " must be on or off");
    }

    /** Several rules may be given in one option, separated by semicolons (outside double quotes). */
    private static void addRules(List<String> into, String v) {
        StringBuilder cur = new StringBuilder();
        boolean inQ = false;
        for (int i = 0; i < v.length(); i++) {
            char ch = v.charAt(i);
            if (ch == '"') inQ = !inQ;
            if (ch == ';' && !inQ) { add(into, cur); continue; }
            cur.append(ch);
        }
        add(into, cur);
    }

    private static void add(List<String> into, StringBuilder cur) {
        String r = cur.toString().trim();
        if (!r.isEmpty()) into.add(r);
        cur.setLength(0);
    }

    boolean hasAction() {
        return freeze != null || gridlines != null || zoom != null || tabcolor != null || autofilter != null
                || autofit || colwidth != null || orientation != null || fitwidth || printtitles != null
                || !colorscale.isEmpty() || !databar.isEmpty() || !highlight.isEmpty() || !dropdown.isEmpty()
                || !comment.isEmpty() || removeComments || groupRows != null || groupCols != null || ungroup
                || protect != null || !unlock.isEmpty() || hideSheet != null
                || summaryAbove != null || summaryLeft != null;
    }

    /** Splits on whitespace, keeping "double quoted" pieces together (quotes removed). */
    static List<String> tokens(String s) {
        List<String> out = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean inQ = false, had = false;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch == '"') { inQ = !inQ; had = true; continue; }
            if (!inQ && Character.isWhitespace(ch)) {
                if (had || cur.length() > 0) { out.add(cur.toString()); cur.setLength(0); had = false; }
                continue;
            }
            cur.append(ch);
        }
        if (had || cur.length() > 0) out.add(cur.toString());
        return out;
    }
}
