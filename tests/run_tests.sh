#!/usr/bin/env bash
# Java-side test suite: builds workbooks with openpyxl and xlsxwriter, runs xlpolish
# on them, and checks the results with openpyxl and a live LibreOffice.
# Needs: a JDK, Python 3 with openpyxl and xlsxwriter, LibreOffice (soffice and its
# Python bridge "uno"). The Stata-side test is ../xlpolish_test.do.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
root="$(dirname "$here")"
classes="$here/work/classes"
rm -rf "$here/work" && mkdir -p "$classes"
javac -d "$classes" "$root/java/stub/com/stata/sfi/SFIToolkit.java"
javac -cp "$root/xlpolish.jar:$classes" -d "$classes" "$here/Run.java"
export XLPOLISH_CP="$root/xlpolish.jar:$classes"
cd "$here/work"
python3 "$here/test_core.py"
python3 "$here/test_features.py"
