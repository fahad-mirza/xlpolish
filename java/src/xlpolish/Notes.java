package xlpolish;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.dom.Element;

/**
 * Cell notes (called "comments" before Excel 365, "Notes" since). Excel stores
 * them in two parts linked from the worksheet: commentsN.xml holds the text and
 * author, and vmlDrawingN.vml holds the yellow pop-up box for each note. When a
 * sheet already has notes, both parts are rebuilt with the old and new notes
 * together, so running xlpolish again adds to (or replaces) notes cell by cell.
 */
final class Notes {

    static final String CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    static final String CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";

    /** One note: author and text, keyed by (row, column) in a sorted map. */
    private static final class Note {
        final String author, text;
        Note(String author, String text) { this.author = author; this.text = text; }
    }

    private Notes() { }

    /**
     * Applies the requested notes to one sheet. adds maps "row,col" to text; empty
     * text deletes that cell's note. removeAll clears every note first.
     * Returns {notes added or changed, notes deleted}.
     */
    static int[] apply(Book book, Book.Sheet sh, Element ws, TreeMap<Long, String> adds, String author, boolean removeAll)
            throws XLPolishException {
        if (book.isStrict()) {
            throw XLPolishException.syntax("this workbook uses the 'strict' Excel format; notes cannot be added to it");
        }
        String[] cRel = book.relOfType(sh.part, "/comments");
        Element legacy = Pkg.first(ws, "legacyDrawing");
        String vmlId = legacy == null ? null : legacy.getAttributeNS(Book.R_NS, "id");
        String[] vRel = null;
        if (vmlId != null && !vmlId.isEmpty()) {
            Element rels = book.relsOf(sh.part);
            for (Element r : Pkg.children(rels, "Relationship")) {
                if (vmlId.equals(r.getAttribute("Id"))) {
                    vRel = new String[] { vmlId, Book.resolve(sh.part.substring(0, sh.part.lastIndexOf('/') + 1), r.getAttribute("Target")) };
                }
            }
        }
        String vmlText = vRel != null && book.pkg.has(vRel[1])
                ? new String(book.pkg.raw(vRel[1]), StandardCharsets.UTF_8) : null;
        if (vmlText != null) {
            Matcher m = Pattern.compile("ObjectType\\s*=\\s*\"(\\w+)\"").matcher(vmlText);
            while (m.find()) {
                if (!"Note".equals(m.group(1))) {
                    throw XLPolishException.syntax("sheet \"" + sh.name + "\" contains form controls (buttons, check boxes) "
                            + "that share the notes drawing; xlpolish cannot edit notes on it");
                }
            }
        }

        // existing notes
        TreeMap<Long, Note> notes = new TreeMap<>();
        if (cRel != null && book.pkg.has(cRel[1]) && !removeAll) {
            Element root = book.pkg.doc(cRel[1]).getDocumentElement();
            List<String> authors = new ArrayList<>();
            Element as = Pkg.first(root, "authors");
            if (as != null) for (Element a : Pkg.children(as, "author")) authors.add(a.getTextContent());
            Element list = Pkg.first(root, "commentList");
            if (list != null) {
                for (Element c : Pkg.children(list, "comment")) {
                    int[] rc = Cells.parseCell(c.getAttribute("ref"));
                    if (rc == null) continue;
                    int ai = Grid.toInt(c.getAttribute("authorId"));
                    Element t = Pkg.first(c, "text");
                    notes.put(key(rc[0], rc[1]), new Note(ai >= 0 && ai < authors.size() ? authors.get(ai) : "",
                            t == null ? "" : Pkg.textOf(t)));
                }
            }
        }
        int added = 0, removed = 0;
        for (Map.Entry<Long, String> e : adds.entrySet()) {
            if (e.getValue().isEmpty()) { if (notes.remove(e.getKey()) != null) removed++; }
            else { notes.put(e.getKey(), new Note(author, e.getValue())); added++; }
        }

        // drop the old parts and links; rebuild if any notes remain
        if (cRel != null) {
            book.removeRel(sh.part, cRel[0]);
            book.pkg.delete(cRel[1]);
            book.removeOverride(cRel[1]);
        }
        if (vRel != null) {
            book.removeRel(sh.part, vRel[0]);
            book.pkg.delete(vRel[1]);
            book.removeOverride(vRel[1]);
        }
        if (legacy != null) ws.removeChild(legacy);
        if (notes.isEmpty()) return new int[] { added, removed };

        int idmap = idmapFor(book, vmlText);
        String cPath = cRel != null ? cRel[1] : book.freeName("xl/comments", ".xml");
        String vPath = vRel != null ? vRel[1] : book.freeName("xl/drawings/vmlDrawing", ".vml");
        book.pkg.putBytes(cPath, commentsXml(notes).getBytes(StandardCharsets.UTF_8));
        book.pkg.putBytes(vPath, vml(notes, idmap).getBytes(StandardCharsets.UTF_8));
        book.ensureOverride(cPath, CT_COMMENTS);
        book.ensureDefault("vml", CT_VML);
        book.addRel(sh.part, Book.REL_COMMENTS, cPath);
        String vid = book.addRel(sh.part, Book.REL_VML, vPath);

        Element ld = Pkg.create(ws, "legacyDrawing");
        ld.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", Book.R_NS);
        ld.setAttributeNS(Book.R_NS, "r:id", vid);
        Pkg.insertOrdered(ws, ld, Pkg.WORKSHEET_ORDER);
        return new int[] { added, removed };
    }

    static long key(int r, int c) { return (long) r * 100000L + c; }
    private static int rowOf(long k) { return (int) (k / 100000L); }
    private static int colOf(long k) { return (int) (k % 100000L); }

    /**
     * Each VML drawing in a workbook needs its own block of shape ids ("idmap").
     * Reuse the sheet's own block if it had one, else take one above every other.
     */
    private static int idmapFor(Book book, String ownVml) throws XLPolishException {
        // any namespace prefix: Excel writes <o:idmap>, other programs e.g. <ns0:idmap>
        Pattern p = Pattern.compile("<(?:[A-Za-z_][\\w.-]*:)?idmap\\b[^>]*\\bdata\\s*=\\s*\"([0-9, ]+)\"");
        if (ownVml != null) {
            Matcher m = p.matcher(ownVml);
            if (m.find()) return Integer.parseInt(m.group(1).split("[, ]+")[0].trim());
        }
        int max = 0;
        for (String n : book.pkg.names()) {
            if (!n.toLowerCase().endsWith(".vml")) continue;
            Matcher m = p.matcher(new String(book.pkg.raw(n), StandardCharsets.UTF_8));
            while (m.find()) for (String d : m.group(1).split("[, ]+")) if (!d.trim().isEmpty()) max = Math.max(max, Integer.parseInt(d.trim()));
        }
        return max + 1;
    }

    private static String commentsXml(TreeMap<Long, Note> notes) {
        List<String> authors = new ArrayList<>();
        for (Note n : notes.values()) if (!authors.contains(n.author)) authors.add(n.author);
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors>");
        for (String a : authors) sb.append("<author>").append(esc(a)).append("</author>");
        sb.append("</authors><commentList>");
        for (Map.Entry<Long, Note> e : notes.entrySet()) {
            sb.append("<comment ref=\"").append(Cells.ref(rowOf(e.getKey()), colOf(e.getKey())))
              .append("\" authorId=\"").append(authors.indexOf(e.getValue().author)).append("\">")
              .append("<text><t xml:space=\"preserve\">").append(esc(e.getValue().text)).append("</t></text></comment>");
        }
        sb.append("</commentList></comments>");
        return sb.toString();
    }

    /** The legacy drawing that shows each note as a yellow box when the cell is hovered. */
    private static String vml(TreeMap<Long, Note> notes, int idmap) {
        StringBuilder sb = new StringBuilder();
        sb.append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" ")
          .append("xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n")
          .append(" <o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"").append(idmap).append("\"/></o:shapelayout>\n")
          .append(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">\n")
          .append("  <v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n")
          .append(" </v:shapetype>\n");
        int i = 0;
        for (Long k : notes.keySet()) {
            int r0 = rowOf(k) - 1, c0 = colOf(k) - 1;           // 0-based
            int top = Math.max(r0 - 1, 0);
            sb.append(" <v:shape id=\"_x0000_s").append(idmap * 1024 + 1 + i++).append("\" type=\"#_x0000_t202\" ")
              .append("style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;")
              .append("z-index:").append(i).append(";visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n")
              .append("  <v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>")
              .append("<v:path o:connecttype=\"none\"/>\n")
              .append("  <v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>\n")
              .append("  <x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>")
              .append("<x:Anchor>").append(c0 + 1).append(", 15, ").append(top).append(", 10, ")
              .append(c0 + 3).append(", 15, ").append(top + 4).append(", 4</x:Anchor>")
              .append("<x:AutoFill>False</x:AutoFill><x:Row>").append(r0).append("</x:Row><x:Column>").append(c0)
              .append("</x:Column></x:ClientData>\n </v:shape>\n");
        }
        sb.append("</xml>");
        return sb.toString();
    }

    private static String esc(String s) {
        StringBuilder sb = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            switch (ch) {
                case '&': sb.append("&amp;"); break;
                case '<': sb.append("&lt;"); break;
                case '>': sb.append("&gt;"); break;
                case '"': sb.append("&quot;"); break;
                default:
                    if (ch < 0x20 && ch != '\n' && ch != '\t' && ch != '\r') break;   // not allowed in XML
                    sb.append(ch);
            }
        }
        return sb.toString();
    }
}
