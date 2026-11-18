#!/bin/sh
# Rebuilds the generated fixtures in .. from these sources, with the applications that write them.
# Run from this directory; needs pandoc, python-docx and LibreOffice (soffice). The fixtures in git
# are the ones to test: rebuilding with other versions of these applications gives other files.
set -e
OUT=..
TMP=$(mktemp -d)
LO() { soffice --headless --convert-to 'docx:MS Word 2007 XML' --outdir "$TMP/out" "$@" >/dev/null 2>&1; }
pandoc report.md -o "$OUT/pandoc-report.docx"
python3 memo.py "$OUT/python-docx-memo.docx"
# LibreOffice Writer saving .docx: its own minutes; the report via ODT; the memo and the Word letter re-saved.
cp minutes.fodt "$TMP/libreoffice-minutes.fodt"
pandoc report.md -o "$TMP/libreoffice-report.odt"
mkdir "$TMP/in"
cp "$OUT/python-docx-memo.docx" "$TMP/in/libreoffice-memo.docx"
cp "$OUT/word-letter-id.docx" "$TMP/in/libreoffice-letter.docx"
LO "$TMP/libreoffice-minutes.fodt" "$TMP/libreoffice-report.odt" "$TMP/in/libreoffice-memo.docx" "$TMP/in/libreoffice-letter.docx"
cp "$TMP"/out/*.docx "$OUT/"
rm -rf "$TMP"
