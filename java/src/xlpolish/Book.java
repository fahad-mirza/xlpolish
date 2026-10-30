package xlpolish;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;

/** The workbook part and what it points to. */
final class Book {

    static final class Sheet {
        final String name;
        final String part;
        final boolean worksheet;
        final int index;          // 0-based position in <sheets>, used as localSheetId
        Sheet(String name, String part, boolean worksheet, int index) {
            this.name = name; this.part = part; this.worksheet = worksheet; this.index = index;
        }
    }

    final Pkg pkg;
    final String wbPath;
    final Document wb;
    final List<Sheet> sheets = new ArrayList<>();
    String sstPath;
    String stylesPath;
    private List<String> sst;

    Book(Pkg pkg) throws XLPolishException {
        this.pkg = pkg;
        String wp = null;
        for (String[] r : rels("_rels/.rels", "").values()) {
            if (r[1].endsWith("/officeDocument")) { wp = r[0]; break; }
        }
        if (wp == null && pkg.has("xl/workbook.xml")) wp = "xl/workbook.xml";
        if (wp == null) throw new XLPolishException(610, "could not find the workbook part; is this an Excel workbook?");
        wbPath = wp;
        wb = pkg.doc(wbPath);

        String dir = dirOf(wbPath);
        Map<String, String[]> wrels = rels(dir + "_rels/" + nameOf(wbPath) + ".rels", dir);
        for (String[] r : wrels.values()) {
            if (r[1].endsWith("/sharedStrings")) sstPath = r[0];
            else if (r[1].endsWith("/styles")) stylesPath = r[0];
        }
        int i = 0;
        for (Element s : Pkg.descendants(wb.getDocumentElement(), "sheet")) {
            String rid = relId(s);
            String[] rel = rid == null ? null : wrels.get(rid);
            if (rel == null) { i++; continue; }
            sheets.add(new Sheet(s.getAttribute("name"), rel[0], rel[1].endsWith("/worksheet"), i));
            i++;
        }
        if (sheets.isEmpty()) throw XLPolishException.syntax("workbook has no sheets");
    }

    List<Sheet> targets(Spec spec) throws XLPolishException {
        List<Sheet> out = new ArrayList<>();
        if (spec.mode == Spec.Mode.ALL) {
            for (Sheet s : sheets) if (s.worksheet) out.add(s);
        } else if (spec.mode == Spec.Mode.FIRST) {
            for (Sheet s : sheets) if (s.worksheet) { out.add(s); break; }
        } else {
            Sheet hit = null;
            for (Sheet s : sheets) if (s.name.equals(spec.sheet)) { hit = s; break; }
            if (hit == null) for (Sheet s : sheets) if (s.name.equalsIgnoreCase(spec.sheet)) { hit = s; break; }
            if (hit == null) {
                List<String> names = new ArrayList<>();
                for (Sheet s : sheets) names.add("\"" + s.name + "\"");
                throw XLPolishException.syntax("sheet \"" + spec.sheet + "\" not found; sheets in this workbook: "
                        + String.join(", ", names));
            }
            if (!hit.worksheet) throw XLPolishException.syntax("sheet \"" + hit.name + "\" is a chart sheet; "
                    + "xlpolish options apply to worksheets only");
            out.add(hit);
        }
        if (out.isEmpty()) throw XLPolishException.syntax("workbook has no worksheets");
        return out;
    }

    /** Shared strings table (empty list if the workbook has none). */
    List<String> sharedStrings() throws XLPolishException {
        if (sst == null) {
            sst = new ArrayList<>();
            if (sstPath != null && pkg.has(sstPath)) {
                for (Element si : Pkg.children(pkg.doc(sstPath).getDocumentElement(), "si")) sst.add(Pkg.textOf(si));
            }
        }
        return sst;
    }

    // ------------------------------------------------- content types / rels

    static final String CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
    static final String REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    static final String R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static final String REL_COMMENTS = R_NS + "/comments";
    static final String REL_VML = R_NS + "/vmlDrawing";

    /** True if the workbook uses the "strict" flavour of the file format (rare). */
    boolean isStrict() {
        String ns = wb.getDocumentElement().getNamespaceURI();
        return ns != null && ns.contains("purl.oclc.org");
    }

    private Element ctRoot() throws XLPolishException {
        if (!pkg.has("[Content_Types].xml")) throw new XLPolishException(610, "workbook has no [Content_Types].xml");
        return pkg.doc("[Content_Types].xml").getDocumentElement();
    }

    void ensureDefault(String ext, String type) throws XLPolishException {
        Element root = ctRoot();
        for (Element d : Pkg.children(root, "Default")) if (ext.equalsIgnoreCase(d.getAttribute("Extension"))) return;
        Element d = root.getOwnerDocument().createElementNS(CT_NS, "Default");
        d.setAttribute("Extension", ext);
        d.setAttribute("ContentType", type);
        Node firstOverride = Pkg.first(root, "Override");
        root.insertBefore(d, firstOverride);
        pkg.markDirty("[Content_Types].xml");
    }

    void ensureOverride(String part, String type) throws XLPolishException {
        Element root = ctRoot();
        String name = "/" + part;
        for (Element o : Pkg.children(root, "Override")) {
            if (name.equalsIgnoreCase(o.getAttribute("PartName"))) {
                if (!type.equals(o.getAttribute("ContentType"))) { o.setAttribute("ContentType", type); pkg.markDirty("[Content_Types].xml"); }
                return;
            }
        }
        Element o = root.getOwnerDocument().createElementNS(CT_NS, "Override");
        o.setAttribute("PartName", name);
        o.setAttribute("ContentType", type);
        root.appendChild(o);
        pkg.markDirty("[Content_Types].xml");
    }

    void removeOverride(String part) throws XLPolishException {
        Element root = ctRoot();
        String name = "/" + part;
        for (Element o : Pkg.children(root, "Override")) {
            if (name.equalsIgnoreCase(o.getAttribute("PartName"))) { root.removeChild(o); pkg.markDirty("[Content_Types].xml"); }
        }
    }

    /** The .rels part that belongs to a part, e.g. xl/worksheets/_rels/sheet1.xml.rels. */
    static String relsPathFor(String part) { return dirOf(part) + "_rels/" + nameOf(part) + ".rels"; }

    /** Relationships of a part, as a live DOM root (created empty if the part has none yet). */
    Element relsOf(String part) throws XLPolishException {
        String rp = relsPathFor(part);
        if (!pkg.has(rp)) {
            Document d = Pkg.parse(("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"" + REL_NS + "\"/>").getBytes(java.nio.charset.StandardCharsets.UTF_8));
            pkg.putDoc(rp, d);
        }
        return pkg.doc(rp).getDocumentElement();
    }

    /** {Id, resolved target} of the first relationship of this type from part, or null. */
    String[] relOfType(String part, String typeSuffix) throws XLPolishException {
        String rp = relsPathFor(part);
        if (!pkg.has(rp)) return null;
        for (Element r : Pkg.children(pkg.doc(rp).getDocumentElement(), "Relationship")) {
            if (r.getAttribute("Type").endsWith(typeSuffix) && !"External".equalsIgnoreCase(r.getAttribute("TargetMode"))) {
                return new String[] { r.getAttribute("Id"), resolve(dirOf(part), r.getAttribute("Target")) };
            }
        }
        return null;
    }

    /** Adds a relationship from part to target; returns its new Id. */
    String addRel(String part, String type, String target) throws XLPolishException {
        Element root = relsOf(part);
        int n = 1;
        java.util.Set<String> used = new java.util.HashSet<>();
        for (Element r : Pkg.children(root, "Relationship")) used.add(r.getAttribute("Id"));
        while (used.contains("rId" + n)) n++;
        Element r = root.getOwnerDocument().createElementNS(REL_NS, "Relationship");
        r.setAttribute("Id", "rId" + n);
        r.setAttribute("Type", type);
        r.setAttribute("Target", relative(dirOf(part), target));
        root.appendChild(r);
        pkg.markDirty(relsPathFor(part));
        return "rId" + n;
    }

    void removeRel(String part, String id) throws XLPolishException {
        String rp = relsPathFor(part);
        if (!pkg.has(rp)) return;
        Element root = pkg.doc(rp).getDocumentElement();
        for (Element r : Pkg.children(root, "Relationship")) if (id.equals(r.getAttribute("Id"))) root.removeChild(r);
        pkg.markDirty(rp);
    }

    /** First unused name like xl/comments{n}.xml. */
    String freeName(String prefix, String suffix) {
        java.util.Set<String> used = new java.util.HashSet<>();
        for (String n : pkg.names()) used.add(n.toLowerCase(java.util.Locale.ROOT));
        int n = 1;
        while (used.contains((prefix + n + suffix).toLowerCase(java.util.Locale.ROOT))) n++;
        return prefix + n + suffix;
    }

    /** Path of target relative to directory fromDir (both package paths without a leading slash). */
    static String relative(String fromDir, String target) {
        String[] a = fromDir.isEmpty() ? new String[0] : fromDir.split("/");
        String[] b = target.split("/");
        int i = 0;
        while (i < a.length && i < b.length - 1 && a[i].equals(b[i])) i++;
        StringBuilder sb = new StringBuilder();
        for (int k = i; k < a.length; k++) sb.append("../");
        for (int k = i; k < b.length; k++) { sb.append(b[k]); if (k < b.length - 1) sb.append('/'); }
        return sb.toString();
    }

    // -------------------------------------------------------------- helpers

    /** Id -> {resolved target, Type}; external targets skipped. */
    private Map<String, String[]> rels(String relsPath, String baseDir) throws XLPolishException {
        Map<String, String[]> out = new HashMap<>();
        if (!pkg.has(relsPath)) return out;
        for (Element r : Pkg.descendants(pkg.doc(relsPath).getDocumentElement(), "Relationship")) {
            if ("External".equalsIgnoreCase(r.getAttribute("TargetMode"))) continue;
            out.put(r.getAttribute("Id"), new String[] { resolve(baseDir, r.getAttribute("Target")), r.getAttribute("Type") });
        }
        return out;
    }

    private static String relId(Element e) {
        NamedNodeMap attrs = e.getAttributes();
        for (int i = 0; i < attrs.getLength(); i++) {
            Attr a = (Attr) attrs.item(i);
            String ln = a.getLocalName() != null ? a.getLocalName() : a.getName();
            if ("id".equals(ln) && a.getNamespaceURI() != null && a.getNamespaceURI().contains("relationships")) {
                return a.getValue();
            }
        }
        return null;
    }

    static String resolve(String baseDir, String target) {
        String p = target.startsWith("/") ? target.substring(1) : baseDir + target;
        Deque<String> parts = new ArrayDeque<>();
        for (String seg : p.split("/")) {
            if (seg.isEmpty() || ".".equals(seg)) continue;
            if ("..".equals(seg)) { if (!parts.isEmpty()) parts.removeLast(); } else parts.addLast(seg);
        }
        return String.join("/", parts);
    }

    private static String dirOf(String part) { int i = part.lastIndexOf('/'); return i < 0 ? "" : part.substring(0, i + 1); }
    private static String nameOf(String part) { int i = part.lastIndexOf('/'); return i < 0 ? part : part.substring(i + 1); }
}
