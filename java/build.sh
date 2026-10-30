#!/usr/bin/env bash
# Rebuilds xlpolish.jar in the repository root from java/src.
# Needs a JDK 9 or newer (for --release 8); the jar runs on Java 8+, i.e. Stata 16+.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
root="$(dirname "$here")"
build="$(mktemp -d)"
trap 'rm -rf "$build"' EXIT
javac --release 8 -Xlint:all -Xlint:-options -d "$build" -sourcepath "$here/stub" "$here"/src/xlpolish/*.java
rm -rf "$build/com"                      # the Stata API stub is for compiling only; Stata supplies the real one
jar cf "$root/xlpolish.jar" -C "$build" xlpolish
echo "built $root/xlpolish.jar"
