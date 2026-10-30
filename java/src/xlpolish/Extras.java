package xlpolish;

import java.time.LocalDate;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeSet;

import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * Features added in 1.3.0: data-validation rules and input prompts, icon sets,
 * top/bottom and above/below-average rules, page headers and footers, hiding
 * rows and columns, and hiding the filter arrow on chosen columns.
 */
final class Extras {

    private Extras() { }

    // ===================================================== data validation

    /** One validate() rule: e.g. "B2:B100 decimal between 0 100 "Enter a score"". */
    static final class Rule {
        String sqref, type, op, f1, f2, message;
    }

    private static final Map<String, String> VTYPES = new HashMap<>();
    private static final Map<String, String> VOPS = new HashMap<>();
    static {
        VTYPES.put("whole", "whole"); VTYPES.put("decimal", "decimal");
        VTYPES.put("date", "date"); VTYPES.put("length", "textLength");
        VOPS.put("between", "between"); VOPS.put("notbetween", "notBetween");
        VOPS.put("==", "equal"); VOPS.put("=", "equal"); VOPS.put("!=", "notEqual"); VOPS.put("~=", "notEqual");
        VOPS.put("<>", "notEqual"); VOPS.put(">", "greaterThan"); VOPS.put(">=", "greaterThanOrEqual");
        VOPS.put("<", "lessThan"); VOPS.put("<=", "lessThanOrEqual");
    }

    static Rule parseRule(String r) throws XLPolishException {
        // pull out a trailing "quoted message" before tokenising
        String s = r.trim(), msg = null;
        if (s.endsWith("\"")) {
            int q = s.lastIndexOf('"', s.length() - 2);
            if (q >= 0) { msg = s.substring(q + 1, s.length() - 1); s = s.substring(0, q).trim(); }
        }
        List<String> t = Spec.tokens(s);
        String usage = "validate() takes: range type rule [\"error message\"], where type is whole, decimal, date or "
                + "length, e.g. validate(B2:B100 decimal between 0 100) or validate(C2:C50 length <= 200)";
        if (t.size() < 4) throw XLPolishException.syntax(usage + "; got: " + r);
        Rule v = new Rule();
        v.sqref = Cells.sqref(t.get(0));
        String type = t.get(1).toLowerCase(Locale.ROOT);
        v.type = VTYPES.get(type);
        if (v.type == null) throw XLPolishException.syntax("validate() type must be whole, decimal, date or length; got \"" + t.get(1) + "\"");
        String op = t.get(2).toLowerCase(Locale.ROOT).replace(" ", "");
        v.op = VOPS.get(op);
        if (v.op == null) throw XLPolishException.syntax("validate() rule must be between, notbetween, ==, !=, >, >=, < or <=; got \"" + t.get(2) + "\"");
        boolean two = v.op.equals("between") || v.op.equals("notBetween");
        if (t.size() != (two ? 5 : 4)) throw XLPolishException.syntax(usage + "; got: " + r);
        v.f1 = value(type, t.get(3));
        if (two) {
            v.f2 = value(type, t.get(4));
            if (Double.parseDouble(v.f1) > Double.parseDouble(v.f2)) {
                throw XLPolishException.syntax("validate(): the lower bound " + t.get(3) + " is above the upper bound " + t.get(4));
            }
        }
        v.message = msg != null ? msg : defaultMessage(type, op, t.get(3), two ? t.get(4) : null);
        if (v.message.length() > 225) throw XLPolishException.syntax("validate() message is longer than Excel's 225-character limit");
        return v;
    }

    private static String value(String type, String tok) throws XLPolishException {
        if (type.equals("date")) {
            try {
                LocalDate d = LocalDate.parse(tok);                   // YYYY-MM-DD
                return Long.toString(ChronoUnit.DAYS.between(LocalDate.of(1899, 12, 30), d));   // Excel's day number
            } catch (DateTimeParseException e) {
                throw XLPolishException.syntax("validate() dates are written YYYY-MM-DD, e.g. 2026-12-31; got \"" + tok + "\"");
            }
        }
        try {
            double d = Double.parseDouble(tok);
            if ((type.equals("whole") || type.equals("length")) && d != Math.rint(d)) {
                throw XLPolishException.syntax("validate() " + type + " limits must be whole numbers; got " + tok);
            }
            if (type.equals("length") && d < 0) throw XLPolishException.syntax("validate() length limits cannot be negative");
            return tok;
        } catch (NumberFormatException e) {
            throw XLPolishException.syntax("validate() limit \"" + tok + "\" is not a number");
        }
    }

    private static String defaultMessage(String type, String op, String a, String b) {
        String what = type.equals("whole") ? "a whole number" : type.equals("decimal") ? "a number"
                : type.equals("date") ? "a date" : "text";
        if (type.equals("length")) {
            switch (op) {
                case "between": return "Enter text between " + a + " and " + b + " characters long.";
                case "notbetween": return "Text length must not be between " + a + " and " + b + " characters.";
                case "<=": return "Enter at most " + a + " characters.";
                case "<": return "Enter fewer than " + a + " characters.";
                case ">=": return "Enter at least " + a + " characters.";
                case ">": return "Enter more than " + a + " characters.";
                default: return "Enter text of the required length.";
            }
        }
        switch (op) {
            case "between": return "Enter " + what + " from " + a + " to " + b + ".";
            case "notbetween": return "Enter " + what + " outside " + a + " to " + b + ".";
            case "==": case "=": return "Enter " + what + " equal to " + a + ".";
            case "!=": case "~=": case "<>": return "Enter " + what + " other than " + a + ".";
            case ">": return "Enter " + what + " greater than " + a + ".";
            case ">=": return "Enter " + what + " of at least " + a + ".";
            case "<": return "Enter " + what + " less than " + a + ".";
            default: return "Enter " + what + " of at most " + a + ".";
        }
    }

    /** One prompt(): cells and the message shown when one is selected. */
    static final class Prompt { String sqref, text; }

    static Prompt parsePrompt(String r) throws XLPolishException {
        String s = r.trim();
        int sp = -1;
        for (int i = 0; i < s.length(); i++) if (Character.isWhitespace(s.charAt(i))) { sp = i; break; }
        if (sp < 0) throw XLPolishException.syntax("prompt() takes a range and then the message, e.g. prompt(B2:B100 \"Enter the 2025 value\"); got: " + r);
        Prompt p = new Prompt();
        p.sqref = Cells.sqref(s.substring(0, sp));
        String text = s.substring(sp + 1).trim();
        if (text.length() >= 2 && text.startsWith("\"") && text.endsWith("\"")) text = text.substring(1, text.length() - 1);
        if (text.isEmpty()) throw XLPolishException.syntax("prompt() message is empty: " + r);
        if (text.length() > 255) throw XLPolishException.syntax("prompt() message is longer than Excel's 255-character limit");
        p.text = text;
        return p;
    }

    /**
     * Finds the validation element that covers exactly sqref (to replace or extend),
     * or throws if another validation overlaps those cells partly: Excel applies one
     * validation per cell, so overlapping rules would conflict.
     */
    static Element sameRange(Element dvs, String sqref, String what) throws XLPolishException {
        List<int[]> mine = Cells.rects(sqref);
        Element exact = null;
        for (Element dv : Pkg.children(dvs, "dataValidation")) {
            String other = dv.getAttribute("sqref");
            List<int[]> theirs;
            try { theirs = Cells.rects(other); } catch (XLPolishException e) { continue; }
            if (normal(other).equals(normal(sqref))) { exact = dv; continue; }
            if (Cells.overlaps(mine, theirs)) {
                throw XLPolishException.syntax(what + " on " + sqref + " overlaps the dropdown or rule already on " + other
                        + "; Excel allows one per cell, so use exactly the same range (to replace it) or a separate one");
            }
        }
        return exact;
    }

    private static String normal(String sq) {
        try {
            StringBuilder b = new StringBuilder();
            for (int[] r : Cells.rects(sq)) b.append(Cells.rectRef(r)).append(' ');
            return b.toString().trim();
        } catch (XLPolishException e) { return sq.trim(); }
    }

    static void addRule(Element ws, Rule v) throws XLPolishException {
        Element dvs = Pkg.ensure(ws, "dataValidations", Pkg.WORKSHEET_ORDER);
        Element old = sameRange(dvs, v.sqref, "validate()");
        Element dv = Pkg.create(dvs, "dataValidation");
        dv.setAttribute("type", v.type);
        if (!v.op.equals("between")) dv.setAttribute("operator", v.op);
        dv.setAttribute("allowBlank", "1");
        dv.setAttribute("showErrorMessage", "1");
        dv.setAttribute("errorTitle", "Invalid entry");
        dv.setAttribute("error", v.message);
        dv.setAttribute("sqref", v.sqref);
        keepPrompt(old, dv);
        Element f1 = Pkg.create(dv, "formula1"); f1.setTextContent(v.f1); dv.appendChild(f1);
        if (v.f2 != null) { Element f2 = Pkg.create(dv, "formula2"); f2.setTextContent(v.f2); dv.appendChild(f2); }
        place(dvs, old, dv);
    }

    /** A dropdown list (from Engine), replacing any rule on exactly the same cells. */
    static void addList(Element ws, String sqref, String formula) throws XLPolishException {
        Element dvs = Pkg.ensure(ws, "dataValidations", Pkg.WORKSHEET_ORDER);
        Element old = sameRange(dvs, sqref, "dropdown()");
        Element dv = Pkg.create(dvs, "dataValidation");
        dv.setAttribute("type", "list");
        dv.setAttribute("allowBlank", "1");
        dv.setAttribute("showErrorMessage", "1");
        dv.setAttribute("sqref", sqref);
        keepPrompt(old, dv);
        Element f = Pkg.create(dv, "formula1"); f.setTextContent(formula); dv.appendChild(f);
        place(dvs, old, dv);
    }

    static void addPrompt(Element ws, Prompt p) throws XLPolishException {
        Element dvs = Pkg.ensure(ws, "dataValidations", Pkg.WORKSHEET_ORDER);
        Element dv = sameRange(dvs, p.sqref, "prompt()");
        if (dv == null) {                                  // a prompt on its own: a validation that allows anything
            dv = Pkg.create(dvs, "dataValidation");
            dv.setAttribute("allowBlank", "1");
            dv.setAttribute("sqref", p.sqref);
            dvs.appendChild(dv);
        }
        dv.setAttribute("showInputMessage", "1");
        dv.setAttribute("prompt", p.text);
        dvs.setAttribute("count", Integer.toString(Pkg.children(dvs, "dataValidation").size()));
    }

    private static void keepPrompt(Element old, Element dv) {
        if (old != null && old.hasAttribute("prompt")) {
            dv.setAttribute("showInputMessage", "1");
            dv.setAttribute("prompt", old.getAttribute("prompt"));
            if (old.hasAttribute("promptTitle")) dv.setAttribute("promptTitle", old.getAttribute("promptTitle"));
        }
    }

    private static void place(Element dvs, Element old, Element dv) {
        if (old != null) { dvs.replaceChild(dv, old); } else { dvs.appendChild(dv); }
        dvs.setAttribute("count", Integer.toString(Pkg.children(dvs, "dataValidation").size()));
    }

    // ================================================= icon sets and rank rules

    static final class IconSet { String sqref, set; int n; boolean percentile, reverse, iconsOnly; }

    private static final Map<String, String> ICONS = new HashMap<>();
    static {
        ICONS.put("arrows", "3Arrows"); ICONS.put("arrowsgray", "3ArrowsGray"); ICONS.put("traffic", "3TrafficLights1");
        ICONS.put("trafficrim", "3TrafficLights2"); ICONS.put("signs", "3Signs"); ICONS.put("symbols", "3Symbols");
        ICONS.put("symbolsplain", "3Symbols2"); ICONS.put("flags", "3Flags");
        ICONS.put("arrows4", "4Arrows"); ICONS.put("arrowsgray4", "4ArrowsGray"); ICONS.put("traffic4", "4TrafficLights");
        ICONS.put("rating4", "4Rating"); ICONS.put("redtoblack4", "4RedToBlack");
        ICONS.put("arrows5", "5Arrows"); ICONS.put("arrowsgray5", "5ArrowsGray"); ICONS.put("rating5", "5Rating");
        ICONS.put("quarters5", "5Quarters");
    }

    static IconSet parseIconSet(String r) throws XLPolishException {
        List<String> t = Spec.tokens(r);
        if (t.size() < 2) throw XLPolishException.syntax("iconset() takes a range and an icon style, e.g. iconset(B2:B75 traffic); got: " + r);
        IconSet ic = new IconSet();
        ic.sqref = Cells.sqref(t.get(0));
        ic.set = ICONS.get(t.get(1).toLowerCase(Locale.ROOT));
        if (ic.set == null) throw XLPolishException.syntax("iconset() style \"" + t.get(1) + "\" is unknown; use one of "
                + "arrows, arrowsgray, traffic, trafficrim, signs, symbols, symbolsplain, flags, arrows4, arrowsgray4, "
                + "traffic4, rating4, redtoblack4, arrows5, arrowsgray5, rating5, quarters5");
        ic.n = ic.set.charAt(0) - '0';
        for (String o : t.subList(2, t.size())) {
            switch (o.toLowerCase(Locale.ROOT)) {
                case "reverse": ic.reverse = true; break;
                case "percentile": ic.percentile = true; break;
                case "iconsonly": ic.iconsOnly = true; break;
                default: throw XLPolishException.syntax("iconset() extra words may be reverse, percentile or iconsonly; got \"" + o + "\"");
            }
        }
        return ic;
    }

    static void addIconSet(Element ws, IconSet ic, int priority) {
        Element rule = cfRule(ws, ic.sqref, "iconSet", priority);
        Element is = Pkg.create(rule, "iconSet");
        is.setAttribute("iconSet", ic.set);
        if (ic.reverse) is.setAttribute("reverse", "1");
        if (ic.iconsOnly) is.setAttribute("showValue", "0");
        rule.appendChild(is);
        for (int k = 0; k < ic.n; k++) {
            Element v = Pkg.create(is, "cfvo");
            v.setAttribute("type", ic.percentile ? "percentile" : "percent");
            v.setAttribute("val", Integer.toString(Math.round(100f * k / ic.n)));
            is.appendChild(v);
        }
    }

    /** top() / bottom(): "B2:B75 10 lightgreen [text]" or "B2:B75 10% ..." */
    static final class Rank { String sqref, fill, font; int n; boolean percent, bottom; }

    static Rank parseRank(String r, boolean bottom) throws XLPolishException {
        String name = bottom ? "bottom()" : "top()";
        List<String> t = Spec.tokens(r);
        if (t.size() < 3 || t.size() > 4) throw XLPolishException.syntax(name + " takes: range N fill-colour [text-colour], "
                + "e.g. " + name.replace("()", "") + "(B2:B75 10 lightgreen) or with a percentage, 10%; got: " + r);
        Rank k = new Rank();
        k.sqref = Cells.sqref(t.get(0));
        String n = t.get(1);
        k.percent = n.endsWith("%");
        if (k.percent) n = n.substring(0, n.length() - 1);
        try { k.n = Integer.parseInt(n); } catch (NumberFormatException e) { k.n = -1; }
        if (k.n < 1 || k.n > (k.percent ? 100 : 1000)) {
            throw XLPolishException.syntax(name + " N must be a whole number from 1 to " + (k.percent ? "100%" : "1000") + "; got " + t.get(1));
        }
        k.fill = Colors.argb(t.get(2));
        k.font = t.size() == 4 ? Colors.argb(t.get(3)) : null;
        k.bottom = bottom;
        return k;
    }

    /** average(): "B2:B75 above lightgreen [text]" */
    static final class Avg { String sqref, fill, font; boolean below; }

    static Avg parseAverage(String r) throws XLPolishException {
        List<String> t = Spec.tokens(r);
        if (t.size() < 3 || t.size() > 4) throw XLPolishException.syntax("average() takes: range above|below fill-colour [text-colour], "
                + "e.g. average(B2:B75 below pink); got: " + r);
        Avg a = new Avg();
        a.sqref = Cells.sqref(t.get(0));
        String w = t.get(1).toLowerCase(Locale.ROOT);
        if (!w.equals("above") && !w.equals("below")) throw XLPolishException.syntax("average() needs above or below; got \"" + t.get(1) + "\"");
        a.below = w.equals("below");
        a.fill = Colors.argb(t.get(2));
        a.font = t.size() == 4 ? Colors.argb(t.get(3)) : null;
        return a;
    }

    static Element cfRule(Element ws, String sqref, String type, int priority) {
        Element cf = Pkg.create(ws, "conditionalFormatting");
        cf.setAttribute("sqref", sqref);
        Element rule = Pkg.create(cf, "cfRule");
        rule.setAttribute("type", type);
        rule.setAttribute("priority", Integer.toString(priority));
        cf.appendChild(rule);
        Pkg.insertOrdered(ws, cf, Pkg.WORKSHEET_ORDER);
        return rule;
    }

    // ===================================================== header / footer

    /**
     * Converts "left | centre | right" text with [page] [pages] [date] [time] [file]
     * [sheet] placeholders to Excel's header code. Returns "" to remove.
     */
    static String headerCode(String raw, String what) throws XLPolishException {
        String s = raw.trim();
        if (s.length() >= 2 && s.startsWith("\"") && s.endsWith("\"")) s = s.substring(1, s.length() - 1);
        if (s.trim().isEmpty()) return "";
        String[] parts = s.split("\\|", -1);
        if (parts.length == 2 || parts.length > 3) {
            throw XLPolishException.syntax(what + "() takes one section (centred) or three separated by | (left | centre | right); got: " + raw);
        }
        String[] tags = parts.length == 1 ? new String[] { "&C" } : new String[] { "&L", "&C", "&R" };
        StringBuilder out = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            String p = parts[i].trim();
            if (p.isEmpty()) continue;
            String code = p.replace("&", "&&")
                    .replace("[pages]", "&N").replace("[page]", "&P").replace("[date]", "&D")
                    .replace("[time]", "&T").replace("[file]", "&F").replace("[sheet]", "&A");
            if (code.matches(".*\\[[a-z]+\\].*")) {
                throw XLPolishException.syntax(what + "() placeholder not recognised in \"" + p
                        + "\"; use [page], [pages], [date], [time], [file] or [sheet]");
            }
            out.append(tags[i]).append(code);
        }
        if (out.length() > 255) throw XLPolishException.syntax(what + "() text is longer than Excel's 255-character limit");
        return out.toString();
    }

    static void setHeaderFooter(Element ws, String header, String footer) {
        Element hf = Pkg.ensure(ws, "headerFooter", Pkg.WORKSHEET_ORDER);
        java.util.List<String> order = java.util.Arrays.asList("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter");
        if (header != null) setPart(hf, "oddHeader", header, order);
        if (footer != null) setPart(hf, "oddFooter", footer, order);
        if (Pkg.children(hf, "oddHeader").isEmpty() && Pkg.children(hf, "oddFooter").isEmpty()
                && Pkg.children(hf, "evenHeader").isEmpty() && Pkg.children(hf, "firstHeader").isEmpty()
                && Pkg.children(hf, "evenFooter").isEmpty() && Pkg.children(hf, "firstFooter").isEmpty()) {
            ws.removeChild(hf);
        }
    }

    private static void setPart(Element hf, String name, String code, List<String> order) {
        for (Element e : Pkg.children(hf, name)) hf.removeChild(e);
        if (code.isEmpty()) return;
        Element e = Pkg.create(hf, name);
        e.setTextContent(code);
        Pkg.insertOrdered(hf, e, order);
    }

    // ===================================================== hide / show

    static List<int[]> parseRowList(String s, String what) throws XLPolishException {
        List<int[]> out = new ArrayList<>();
        for (String tok : s.trim().split("[\\s,]+")) {
            if (tok.isEmpty()) continue;
            String u = tok.replace("$", "");
            if (!u.matches("[0-9]+(:[0-9]+)?")) throw XLPolishException.syntax(what + "() takes rows such as 5 or 5:20; got \"" + tok + "\"");
            String[] p = u.split(":");
            long a = Long.parseLong(p[0]), b = Long.parseLong(p[p.length - 1]);
            if (a < 1 || a > b || b > Cells.MAX_ROWS) throw XLPolishException.syntax("invalid rows in " + what + "(): " + tok);
            out.add(new int[] { (int) a, (int) b });
        }
        if (out.isEmpty()) throw XLPolishException.syntax(what + "() is empty");
        return out;
    }

    static List<int[]> parseColList(String s, String what) throws XLPolishException {
        List<int[]> out = new ArrayList<>();
        for (String tok : s.trim().split("[\\s,]+")) {
            if (tok.isEmpty()) continue;
            String u = tok.replace("$", "").toUpperCase(Locale.ROOT);
            if (!u.matches("[A-Z]{1,3}(:[A-Z]{1,3})?")) throw XLPolishException.syntax(what + "() takes columns such as E or C:F; got \"" + tok + "\"");
            String[] p = u.split(":");
            int a = Cells.colNum(p[0]), b = Cells.colNum(p[p.length - 1]);
            if (a > b || b > Cells.MAX_COLS) throw XLPolishException.syntax("invalid columns in " + what + "(): " + tok);
            out.add(new int[] { a, b });
        }
        if (out.isEmpty()) throw XLPolishException.syntax(what + "() is empty");
        return out;
    }

    static void rowsHidden(Element ws, List<int[]> ranges, boolean hide) {
        Grid g = new Grid(ws);
        for (int[] rg : ranges) {
            for (int r = rg[0]; r <= rg[1]; r++) {
                if (hide) g.row(r).setAttribute("hidden", "1");
                else { Element row = g.rows().get(r); if (row != null) row.removeAttribute("hidden"); }
            }
        }
    }

    static void colsHidden(Element ws, List<int[]> ranges, final boolean hide) {
        for (int[] rg : ranges) {
            for (int c = rg[0]; c <= rg[1]; c++) {
                if (!hide && !covered(ws, c)) continue;          // nothing to show
                Grid.modifyCol(ws, c, col -> { if (hide) col.setAttribute("hidden", "1"); else col.removeAttribute("hidden"); });
            }
        }
    }

    private static boolean covered(Element ws, int c) {
        for (Element col : Grid.cols(ws)) {
            if (Grid.toInt(col.getAttribute("min")) <= c && c <= Grid.toInt(col.getAttribute("max"))) return true;
        }
        return false;
    }

    // ===================================================== filter arrows

    /** Hides the filter dropdown arrow on the given columns of the sheet's filter. */
    static void hideFilterArrows(Element ws, List<int[]> cols, String sheetName) throws XLPolishException {
        Element af = Pkg.first(ws, "autoFilter");
        if (af == null) {
            throw XLPolishException.syntax("nofilterarrow() needs a filter on sheet \"" + sheetName
                    + "\"; add autofilter or filterrange() in the same call");
        }
        int[] rect = Cells.parseRect(af.getAttribute("ref"));
        TreeSet<Integer> want = new TreeSet<>();
        for (int[] rg : cols) {
            for (int c = rg[0]; c <= rg[1]; c++) {
                if (c < rect[1] || c > rect[3]) {
                    throw XLPolishException.syntax("nofilterarrow(): column " + Cells.colName(c) + " is outside the filter range "
                            + Cells.rectRef(rect));
                }
                want.add(c - rect[1]);
            }
        }
        for (int id : want) {
            Element hit = null;
            for (Element fc : Pkg.children(af, "filterColumn")) if (Grid.toInt(fc.getAttribute("colId")) == id && fc.hasAttribute("colId")) hit = fc;
            if (hit == null) {
                hit = Pkg.create(af, "filterColumn");
                hit.setAttribute("colId", Integer.toString(id));
                Node before = null;                         // keep filterColumn in colId order, before sortState/extLst
                for (Node n = af.getFirstChild(); n != null; n = n.getNextSibling()) {
                    if (n.getNodeType() != Node.ELEMENT_NODE) continue;
                    if (!"filterColumn".equals(n.getLocalName()) || Grid.toInt(((Element) n).getAttribute("colId")) > id) { before = n; break; }
                }
                af.insertBefore(hit, before);
            }
            hit.setAttribute("hiddenButton", "1");
        }
    }
}
