package xlpolish;

/** An error with the Stata return code to exit with. */
final class XLPolishException extends Exception {
    private static final long serialVersionUID = 1L;
    final int rc;
    XLPolishException(int rc, String msg) { super(msg); this.rc = rc; }
    static XLPolishException syntax(String msg) { return new XLPolishException(198, msg); }
}
