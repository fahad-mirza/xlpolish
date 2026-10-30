package com.stata.sfi;
/** Compile-time stand-in for Stata's SFIToolkit (not shipped in the jar). */
public final class SFIToolkit {
    public static void displayln(String s) { System.out.println(s); }
    public static void errorln(String s) { System.out.println("ERR: " + s); }
}
