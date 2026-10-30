package xlpolish;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * An .xlsx/.xlsm package opened for editing. Parts are parsed on demand and
 * cached; only parts marked dirty are re-serialised on save. Every other zip
 * entry is copied through byte for byte. The new file is written to a temp file
 * in the same folder and then swapped in, so a failure never leaves a partial file.
 */
final class Pkg implements Closeable {

    final Path file;
    private ZipFile zip;
    private final Map<String, String> lower = new HashMap<>();
    private final Map<String, Document> docs = new HashMap<>();
    private final Set<String> dirty = new LinkedHashSet<>();
    /** Raw-byte replacements for existing parts, or contents of new parts. */
    private final Map<String, byte[]> bytes = new LinkedHashMap<>();
    /** Parts that did not exist in the original file (written after the existing ones). */
    private final Set<String> created = new LinkedHashSet<>();
    private final Set<String> deleted = new LinkedHashSet<>();

    Pkg(Path file) throws XLPolishException {
        this.file = file;
        if (!Files.isRegularFile(file)) throw new XLPolishException(601, "file " + file + " not found");
        try {
            zip = new ZipFile(file.toFile());
        } catch (IOException e) {
            throw new XLPolishException(610, "file " + file
                    + " is not a valid .xlsx/.xlsm workbook (old .xls files are not supported)");
        }
        for (Enumeration<? extends ZipEntry> e = zip.entries(); e.hasMoreElements(); ) {
            String n = e.nextElement().getName();
            lower.put(n.toLowerCase(Locale.ROOT), n);
        }
    }

    boolean has(String part) {
        if (part == null) return false;
        String a = lower.get(part.toLowerCase(Locale.ROOT));
        return a != null && !deleted.contains(a);
    }

    /** Names of all parts currently in the package. */
    List<String> names() {
        List<String> out = new ArrayList<>();
        for (String a : lower.values()) if (!deleted.contains(a)) out.add(a);
        return out;
    }

    /** Adds a new part, or replaces an existing one, with raw bytes. */
    void putBytes(String part, byte[] data) {
        String a = lower.get(part.toLowerCase(Locale.ROOT));
        if (a == null) { a = part; lower.put(part.toLowerCase(Locale.ROOT), part); created.add(part); }
        deleted.remove(a);
        docs.remove(a);
        dirty.remove(a);
        bytes.put(a, data);
    }

    /** Adds a new XML part held as a DOM document. */
    void putDoc(String part, Document d) {
        String a = lower.get(part.toLowerCase(Locale.ROOT));
        if (a == null) { a = part; lower.put(part.toLowerCase(Locale.ROOT), part); created.add(part); }
        deleted.remove(a);
        bytes.remove(a);
        docs.put(a, d);
        dirty.add(a);
    }

    /** Removes a part from the package. */
    void delete(String part) {
        String a = lower.get(part.toLowerCase(Locale.ROOT));
        if (a == null) return;
        deleted.add(a);
        dirty.remove(a);
        bytes.remove(a);
        docs.remove(a);
    }

    /** Raw bytes of a part (its pending replacement if there is one). */
    byte[] raw(String part) throws XLPolishException {
        String a = actual(part);
        byte[] b = bytes.get(a);
        return b != null ? b : read(a);
    }

    private String actual(String part) throws XLPolishException {
        String a = part == null ? null : lower.get(part.toLowerCase(Locale.ROOT));
        if (a == null || deleted.contains(a)) throw new XLPolishException(610, "part " + part + " is missing from the workbook");
        return a;
    }

    Document doc(String part) throws XLPolishException {
        String a = actual(part);
        Document d = docs.get(a);
        if (d == null) {
            byte[] b = bytes.get(a);
            d = parse(b != null ? b : read(a));
            bytes.remove(a);
            docs.put(a, d);
            if (b != null) dirty.add(a);
        }
        return d;
    }

    void markDirty(String part) throws XLPolishException { dirty.add(actual(part)); }

    boolean anyDirty() { return !dirty.isEmpty() || !bytes.isEmpty() || !deleted.isEmpty(); }

    private byte[] read(String actualName) throws XLPolishException {
        ZipEntry ze = zip.getEntry(actualName);
        if (ze == null) throw new XLPolishException(610, "part " + actualName + " is missing from the workbook");
        try (InputStream in = zip.getInputStream(ze)) {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            byte[] buf = new byte[1 << 16];
            int n;
            while ((n = in.read(buf)) > 0) bos.write(buf, 0, n);
            return bos.toByteArray();
        } catch (IOException e) {
            throw new XLPolishException(603, "could not read " + actualName + ": " + e.getMessage());
        }
    }

    /** Writes dirty parts and atomically replaces the file. */
    void save() throws XLPolishException {
        if (!anyDirty()) return;
        Map<String, byte[]> rep = new HashMap<>(bytes);
        for (String p : dirty) rep.put(p, serialize(docs.get(p)));

        Path dir = file.toAbsolutePath().getParent();
        Path tmp;
        try {
            tmp = Files.createTempFile(dir, ".xlpolish-", ".tmp");
        } catch (IOException e) {
            throw new XLPolishException(603, "could not create a temporary file in " + dir + ": " + e.getMessage());
        }
        try {
            try (ZipOutputStream zos = new ZipOutputStream(Files.newOutputStream(tmp))) {
                byte[] buf = new byte[1 << 16];
                for (Enumeration<? extends ZipEntry> e = zip.entries(); e.hasMoreElements(); ) {
                    ZipEntry src = e.nextElement();
                    if (deleted.contains(src.getName())) continue;
                    ZipEntry dst = new ZipEntry(src.getName());
                    if (src.getTime() != -1) dst.setTime(src.getTime());
                    zos.putNextEntry(dst);
                    byte[] r = rep.get(src.getName());
                    if (r != null) {
                        zos.write(r);
                    } else if (!src.isDirectory()) {
                        try (InputStream in = zip.getInputStream(src)) {
                            int n;
                            while ((n = in.read(buf)) > 0) zos.write(buf, 0, n);
                        }
                    }
                    zos.closeEntry();
                }
                for (String p : created) {
                    if (deleted.contains(p)) continue;
                    byte[] r = rep.get(p);
                    if (r == null) continue;
                    zos.putNextEntry(new ZipEntry(p));
                    zos.write(r);
                    zos.closeEntry();
                }
            }
            zip.close();
            zip = null;
            try {
                Files.move(tmp, file, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException e) {
                Files.move(tmp, file, StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException e) {
            try { Files.deleteIfExists(tmp); } catch (IOException ignored) { }
            throw new XLPolishException(603, "could not save " + file
                    + " (if it is open in Excel, close it and try again): " + e.getMessage());
        }
    }

    @Override public void close() {
        if (zip != null) { try { zip.close(); } catch (IOException ignored) { } zip = null; }
    }

    // ------------------------------------------------------------------ XML

    static Document parse(byte[] xml) throws XLPolishException {
        try {
            DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
            f.setNamespaceAware(true);
            f.setExpandEntityReferences(false);
            f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            f.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            return f.newDocumentBuilder().parse(new ByteArrayInputStream(xml));
        } catch (Exception e) {
            throw new XLPolishException(610, "could not parse workbook XML: " + e.getMessage());
        }
    }

    static byte[] serialize(Document doc) throws XLPolishException {
        try {
            Transformer t = TransformerFactory.newInstance().newTransformer();
            t.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            t.setOutputProperty(OutputKeys.STANDALONE, "yes");
            t.setOutputProperty(OutputKeys.INDENT, "no");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            t.transform(new DOMSource(doc), new StreamResult(bos));
            return bos.toByteArray();
        } catch (Exception e) {
            throw new XLPolishException(610, "could not write workbook XML: " + e.getMessage());
        }
    }

    // ------------------------------------------------- element order (schema)

    /** Child order of &lt;worksheet&gt; (CT_Worksheet). Excel rejects files that break it. */
    static final List<String> WORKSHEET_ORDER = Arrays.asList(
            "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
            "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
            "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
            "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
            "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems",
            "tableParts", "extLst");

    /** Child order of &lt;sheetPr&gt;. */
    static final List<String> SHEETPR_ORDER = Arrays.asList("tabColor", "outlinePr", "pageSetUpPr");

    /** Child order of &lt;sheetView&gt;. */
    static final List<String> SHEETVIEW_ORDER = Arrays.asList("pane", "selection", "pivotSelection", "extLst");

    /** Child order of &lt;workbook&gt; (CT_Workbook). */
    static final List<String> WORKBOOK_ORDER = Arrays.asList(
            "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
            "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews",
            "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr",
            "webPublishObjects", "extLst");

    /** Child order of &lt;xf&gt; in cellXfs. */
    static final List<String> XF_ORDER = Arrays.asList("alignment", "protection", "extLst");

    /** Child order of &lt;styleSheet&gt;. */
    static final List<String> STYLESHEET_ORDER = Arrays.asList(
            "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
            "tableStyles", "colors", "extLst");

    /**
     * Inserts child into parent at the position the schema requires: before the
     * first known sibling that must come later. Unknown siblings (markup-compatibility
     * wrappers, extensions) are left where they are.
     */
    static void insertOrdered(Element parent, Element child, List<String> order) {
        int r = order.indexOf(child.getLocalName());
        for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n.getNodeType() != Node.ELEMENT_NODE) continue;
            int rk = order.indexOf(n.getLocalName());
            if (rk >= 0 && rk > r) { parent.insertBefore(child, n); return; }
        }
        parent.appendChild(child);
    }

    /** Returns the first child with this local name, creating it in schema order if absent. */
    static Element ensure(Element parent, String local, List<String> order) {
        Element e = first(parent, local);
        if (e == null) {
            e = create(parent, local);
            insertOrdered(parent, e, order);
        }
        return e;
    }

    /** New element in the same namespace and with the same prefix as parent. */
    static Element create(Element parent, String local) {
        String ns = parent.getNamespaceURI();
        String pfx = parent.getPrefix();
        return parent.getOwnerDocument().createElementNS(ns, pfx == null || pfx.isEmpty() ? local : pfx + ":" + local);
    }

    static Element first(Element parent, String local) {
        List<Element> c = children(parent, local);
        return c.isEmpty() ? null : c.get(0);
    }

    /** Direct children with this local name in the parent's namespace. */
    static List<Element> children(Element parent, String local) {
        String ns = parent.getNamespaceURI();
        List<Element> out = new ArrayList<>();
        for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n.getNodeType() == Node.ELEMENT_NODE && local.equals(n.getLocalName())
                    && (ns == null ? n.getNamespaceURI() == null : ns.equals(n.getNamespaceURI()))) {
                out.add((Element) n);
            }
        }
        return out;
    }

    /** All descendants with this local name, in document order, any namespace. */
    static List<Element> descendants(Element root, String local) {
        List<Element> out = new ArrayList<>();
        collect(root, local, out);
        return out;
    }

    private static void collect(Node n, String local, List<Element> out) {
        for (Node k = n.getFirstChild(); k != null; k = k.getNextSibling()) {
            if (k.getNodeType() != Node.ELEMENT_NODE) continue;
            if (local.equals(k.getLocalName())) out.add((Element) k);
            collect(k, local, out);
        }
    }

    /** Concatenated text of all &lt;t&gt; descendants, skipping phonetic runs (&lt;rPh&gt;). */
    static String textOf(Element e) {
        StringBuilder sb = new StringBuilder();
        textInto(e, sb);
        return sb.toString();
    }

    private static void textInto(Node n, StringBuilder sb) {
        for (Node k = n.getFirstChild(); k != null; k = k.getNextSibling()) {
            if (k.getNodeType() != Node.ELEMENT_NODE) continue;
            String ln = k.getLocalName();
            if ("rPh".equals(ln)) continue;
            if ("t".equals(ln)) sb.append(k.getTextContent());
            else textInto(k, sb);
        }
    }
}
