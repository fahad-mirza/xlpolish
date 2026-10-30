package xlpolish;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Colour parsing. Accepts the same colour names as putexcel (CSS names),
 * hex "1F4E79" or "#1F4E79", and "R G B" triples such as "31 78 121".
 * Returns Excel ARGB ("FF1F4E79").
 */
final class Colors {
    private static final Map<String, String> M = new HashMap<>();
    private static final Pattern HEX = Pattern.compile("^#?([0-9A-Fa-f]{6})$");
    private static final Pattern RGB = Pattern.compile("^(\\d{1,3})[\\s,]+(\\d{1,3})[\\s,]+(\\d{1,3})$");

    static {
        M.put("aliceblue", "F0F8FF");
        M.put("antiquewhite", "FAEBD7");
        M.put("aqua", "00FFFF");
        M.put("aquamarine", "7FFFD4");
        M.put("azure", "F0FFFF");
        M.put("beige", "F5F5DC");
        M.put("bisque", "FFE4C4");
        M.put("black", "000000");
        M.put("blanchedalmond", "FFEBCD");
        M.put("blue", "0000FF");
        M.put("blueviolet", "8A2BE2");
        M.put("brown", "A52A2A");
        M.put("burlywood", "DEB887");
        M.put("cadetblue", "5F9EA0");
        M.put("chartreuse", "7FFF00");
        M.put("chocolate", "D2691E");
        M.put("coral", "FF7F50");
        M.put("cornflowerblue", "6495ED");
        M.put("cornsilk", "FFF8DC");
        M.put("crimson", "DC143C");
        M.put("cyan", "00FFFF");
        M.put("darkblue", "00008B");
        M.put("darkcyan", "008B8B");
        M.put("darkgoldenrod", "B8860B");
        M.put("darkgray", "A9A9A9");
        M.put("darkgreen", "006400");
        M.put("darkkhaki", "BDB76B");
        M.put("darkmagenta", "8B008B");
        M.put("darkolivegreen", "556B2F");
        M.put("darkorange", "FF8C00");
        M.put("darkorchid", "9932CC");
        M.put("darkred", "8B0000");
        M.put("darksalmon", "E9967A");
        M.put("darkseagreen", "8FBC8F");
        M.put("darkslateblue", "483D8B");
        M.put("darkslategray", "2F4F4F");
        M.put("darkturquoise", "00CED1");
        M.put("darkviolet", "9400D3");
        M.put("deeppink", "FF1493");
        M.put("deepskyblue", "00BFFF");
        M.put("dimgray", "696969");
        M.put("dodgerblue", "1E90FF");
        M.put("firebrick", "B22222");
        M.put("floralwhite", "FFFAF0");
        M.put("forestgreen", "228B22");
        M.put("fuchsia", "FF00FF");
        M.put("gainsboro", "DCDCDC");
        M.put("ghostwhite", "F8F8FF");
        M.put("gold", "FFD700");
        M.put("goldenrod", "DAA520");
        M.put("gray", "808080");
        M.put("green", "008000");
        M.put("greenyellow", "ADFF2F");
        M.put("honeydew", "F0FFF0");
        M.put("hotpink", "FF69B4");
        M.put("indianred", "CD5C5C");
        M.put("indigo", "4B0082");
        M.put("ivory", "FFFFF0");
        M.put("khaki", "F0E68C");
        M.put("lavender", "E6E6FA");
        M.put("lavenderblush", "FFF0F5");
        M.put("lawngreen", "7CFC00");
        M.put("lemonchiffon", "FFFACD");
        M.put("lightblue", "ADD8E6");
        M.put("lightcoral", "F08080");
        M.put("lightcyan", "E0FFFF");
        M.put("lightgoldenrodyellow", "FAFAD2");
        M.put("lightgray", "D3D3D3");
        M.put("lightgreen", "90EE90");
        M.put("lightpink", "FFB6C1");
        M.put("lightsalmon", "FFA07A");
        M.put("lightseagreen", "20B2AA");
        M.put("lightskyblue", "87CEFA");
        M.put("lightslategray", "778899");
        M.put("lightsteelblue", "B0C4DE");
        M.put("lightyellow", "FFFFE0");
        M.put("lime", "00FF00");
        M.put("limegreen", "32CD32");
        M.put("linen", "FAF0E6");
        M.put("magenta", "FF00FF");
        M.put("maroon", "800000");
        M.put("mediumaquamarine", "66CDAA");
        M.put("mediumblue", "0000CD");
        M.put("mediumorchid", "BA55D3");
        M.put("mediumpurple", "9370DB");
        M.put("mediumseagreen", "3CB371");
        M.put("mediumslateblue", "7B68EE");
        M.put("mediumspringgreen", "00FA9A");
        M.put("mediumturquoise", "48D1CC");
        M.put("mediumvioletred", "C71585");
        M.put("midnightblue", "191970");
        M.put("mintcream", "F5FFFA");
        M.put("mistyrose", "FFE4E1");
        M.put("moccasin", "FFE4B5");
        M.put("navajowhite", "FFDEAD");
        M.put("navy", "000080");
        M.put("oldlace", "FDF5E6");
        M.put("olive", "808000");
        M.put("olivedrab", "6B8E23");
        M.put("orange", "FFA500");
        M.put("orangered", "FF4500");
        M.put("orchid", "DA70D6");
        M.put("palegoldenrod", "EEE8AA");
        M.put("palegreen", "98FB98");
        M.put("paleturquoise", "AFEEEE");
        M.put("palevioletred", "DB7093");
        M.put("papayawhip", "FFEFD5");
        M.put("peachpuff", "FFDAB9");
        M.put("peru", "CD853F");
        M.put("pink", "FFC0CB");
        M.put("plum", "DDA0DD");
        M.put("powderblue", "B0E0E6");
        M.put("powerblue", "B0E0E6");
        M.put("purple", "800080");
        M.put("red", "FF0000");
        M.put("rosybrown", "BC8F8F");
        M.put("royalblue", "4169E1");
        M.put("saddlebrown", "8B4513");
        M.put("salmon", "FA8072");
        M.put("sandybrown", "F4A460");
        M.put("seagreen", "2E8B57");
        M.put("seashell", "FFF5EE");
        M.put("sienna", "A0522D");
        M.put("silver", "C0C0C0");
        M.put("skyblue", "87CEEB");
        M.put("slateblue", "6A5ACD");
        M.put("slategray", "708090");
        M.put("snow", "FFFAFA");
        M.put("springgreen", "00FF7F");
        M.put("steelblue", "4682B4");
        M.put("tan", "D2B48C");
        M.put("teal", "008080");
        M.put("thistle", "D8BFD8");
        M.put("tomato", "FF6347");
        M.put("turquoise", "40E0D0");
        M.put("violet", "EE82EE");
        M.put("wheat", "F5DEB3");
        M.put("white", "FFFFFF");
        M.put("whitesmoke", "F5F5F5");
        M.put("yellow", "FFFF00");
        M.put("yellowgreen", "9ACD32");
        M.put("grey", "808080");
        M.put("darkgrey", "A9A9A9");
        M.put("lightgrey", "D3D3D3");
        M.put("dimgrey", "696969");
        M.put("slategrey", "708090");
        M.put("darkslategrey", "2F4F4F");
        M.put("lightslategrey", "778899");
    }

    private Colors() { }

    static String argb(String spec) throws XLPolishException {
        String s = spec == null ? "" : spec.trim();
        Matcher m = HEX.matcher(s);
        if (m.matches()) return "FF" + m.group(1).toUpperCase(Locale.ROOT);
        m = RGB.matcher(s);
        if (m.matches()) {
            int r = Integer.parseInt(m.group(1)), g = Integer.parseInt(m.group(2)), b = Integer.parseInt(m.group(3));
            if (r <= 255 && g <= 255 && b <= 255) return String.format("FF%02X%02X%02X", r, g, b);
        }
        String h = M.get(s.toLowerCase(Locale.ROOT).replace(" ", ""));
        if (h != null) return "FF" + h;
        throw XLPolishException.syntax("\"" + spec + "\" is not a colour; use a name such as lightgreen, "
                + "a hex code such as C6EFCE, or \"R G B\" such as \"198 239 206\"");
    }

    static boolean isColor(String spec) {
        try { argb(spec); return true; } catch (XLPolishException e) { return false; }
    }
}
