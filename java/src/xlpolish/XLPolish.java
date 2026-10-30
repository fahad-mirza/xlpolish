package xlpolish;

import java.nio.file.Paths;
import java.util.List;

import com.stata.sfi.SFIToolkit;

/**
 * Stata entry point:
 *   javacall xlpolish.XLPolish run, jars(xlpolish.jar) args(optionfile)
 * The option file is written by xlpolish.ado (see Spec). Returns 0 on success,
 * otherwise a Stata return code.
 */
public final class XLPolish {

    private XLPolish() { }

    public static int run(String[] args) {
        try {
            if (args == null || args.length != 1) {
                SFIToolkit.errorln("xlpolish: internal error, expected one argument (option file)");
                return 198;
            }
            Spec spec = Spec.read(Paths.get(args[0]));
            List<String> lines = new Engine(spec).run();
            for (String l : lines) SFIToolkit.displayln(l);
            return 0;
        } catch (XLPolishException e) {
            SFIToolkit.errorln("xlpolish: " + e.getMessage());
            return e.rc;
        } catch (Throwable t) {
            SFIToolkit.errorln("xlpolish: unexpected error: " + t);
            return 198;
        }
    }
}
