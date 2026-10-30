# Java-side tests

`run_tests.sh` builds workbooks with openpyxl and xlsxwriter (and from LibreOffice
output), runs xlpolish on them through its real Stata entry point, and checks the
results three ways: by reading the XML, with openpyxl, and with a live LibreOffice
(which, for example, must refuse a wrong sheet password and accept the right one).

Requirements: a JDK, Python 3 with `openpyxl` and `xlsxwriter`, and LibreOffice
with its Python bridge (`uno`). Run from anywhere:

    ./tests/run_tests.sh

Each suite ends with `FAILURES: 0`. Working files go to `tests/work/` (git-ignored).

The Stata-side test is `../xlpolish_test.do`.
