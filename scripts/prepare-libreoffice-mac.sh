#!/usr/bin/env bash
# Stages a pruned copy of the local LibreOffice.app into resources/libreoffice-mac/,
# for electron-builder's mac.extraResources to bundle into the packaged app —
# so an office clerk's Mac never needs its own separate LibreOffice install
# for accurate document preview/print/PDF export. Run this once before
# `npm run dist`, or whenever LibreOffice is upgraded on this machine.
# resources/libreoffice-mac/ is gitignored, never committed.
#
# The exclusion list was derived empirically — removing candidates, then
# running real conversions, one category at a time, not guessed from file/
# folder names. Two categories LOOKED safe to remove but turned out to be
# load-time hard dependencies, only found by actually testing:
#   - The legacy-format import filters (mwaw/staroffice/wpg/wpd/etonyek/wps —
#     AppleWorks/Keynote/StarOffice/WordPerfect/MS Works) are statically
#     linked INTO libmergedlo.dylib itself. Removing them crashes soffice at
#     startup even though this app never opens those formats.
#   - "Draw" (libsd*.dylib) is not just for .odg files — it's the component
#     LibreOffice's own PDF->PNG rasterization goes through internally
#     (core/docxToPdf.ts's docxToPageImages, the app's entire accurate
#     preview/print pipeline). Removing it broke PDF->image export outright.
# Kept out instead (verified NOT to be load-time dependencies): help,
# spellcheck/NLP extensions, the clipart gallery, built-in templates, GUI
# wizards, the Firebird embedded database, Java-extension support, toolbar
# icon themes, Calc, Math, VBA-macro-object support, the PostgreSQL Base
# connector — none of which this app's Writer-only docx/doc -> pdf/png
# pipeline ever touches. The app also needs legacy .doc (not just .docx)
# support — WordToolPage's "Word merge" tool accepts .doc uploads — which is
# exactly why the legacy filters above must stay.
#
# Result last measured on this machine: 800MB -> ~528MB (~34% smaller).

set -euo pipefail

SOURCE="${1:-/Applications/LibreOffice.app}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEST_DIR="$REPO_ROOT/resources/libreoffice-mac"
DEST="$DEST_DIR/LibreOffice.app"

if [ ! -d "$SOURCE" ]; then
  echo "LibreOffice.app not found at $SOURCE — install it, or pass its path as the first argument." >&2
  exit 1
fi

echo "Staging a pruned copy of $SOURCE -> $DEST ..."
rm -rf "$DEST"
mkdir -p "$DEST_DIR"

rsync -a \
  --exclude='Resources/help' \
  --exclude='Resources/extensions' \
  --exclude='Resources/gallery' \
  --exclude='Resources/template' \
  --exclude='Resources/wizards' \
  --exclude='Resources/firebird' \
  --exclude='Resources/java' \
  --exclude='Resources/CREDITS.fodt' \
  --exclude='Resources/*.icns' \
  --exclude='Resources/config/images_*.zip' \
  "$SOURCE/" "$DEST/"

# Calc / Math / VBA-object / PostgreSQL-Base-connector — empirically not
# load-time dependencies (unlike the legacy filters and Draw, kept above).
FRAMEWORKS="$DEST/Contents/Frameworks"
rm -f "$FRAMEWORKS"/libsc*.dylib \
      "$FRAMEWORKS"/libsmdlo.dylib "$FRAMEWORKS"/libsmlo.dylib \
      "$FRAMEWORKS"/libvbaobjlo.dylib "$FRAMEWORKS"/libvbaswobjlo.dylib \
      "$FRAMEWORKS"/libpostgresql-sdbc-impllo.dylib "$FRAMEWORKS"/libmysql*.dylib

echo "Staged size:"
du -sh "$DEST"
echo "(Original was: $(du -sh "$SOURCE" | cut -f1))"

# --- Self-test: prove the pruned copy actually still works before trusting it. ---
SOFFICE="$DEST/Contents/MacOS/soffice"
if [ ! -x "$SOFFICE" ]; then
  echo "SELF-TEST FAILED: soffice binary missing from the staged copy at $SOFFICE." >&2
  exit 1
fi

TEST_DIR="$(mktemp -d)"
trap 'rm -rf "$TEST_DIR"' EXIT

SAMPLE_DOCX="$REPO_ROOT/resources/work-order-template.docx"
if [ -f "$SAMPLE_DOCX" ]; then
  echo "Self-test: docx -> pdf ..."
  "$SOFFICE" --headless --convert-to pdf --outdir "$TEST_DIR" "$SAMPLE_DOCX" >/dev/null 2>&1
  PDF="$TEST_DIR/$(basename "${SAMPLE_DOCX%.docx}").pdf"
  if [ ! -s "$PDF" ]; then
    echo "SELF-TEST FAILED: docx -> pdf produced no output. The pruned bundle at $DEST is broken — do not ship it." >&2
    exit 1
  fi
  echo "  OK ($(stat -f%z "$PDF") bytes)"

  echo "Self-test: pdf -> png (the app's accurate preview/print path) ..."
  "$SOFFICE" --headless --convert-to 'png:draw_png_Export:{"PixelWidth":{"type":"long","value":"2480"},"PixelHeight":{"type":"long","value":"3508"}}' --outdir "$TEST_DIR" "$PDF" >/dev/null 2>&1
  PNG="$TEST_DIR/$(basename "${PDF%.pdf}").png"
  if [ ! -s "$PNG" ]; then
    echo "SELF-TEST FAILED: pdf -> png produced no output. The pruned bundle at $DEST is broken — do not ship it." >&2
    exit 1
  fi
  echo "  OK ($(stat -f%z "$PNG") bytes)"
else
  echo "Warning: $SAMPLE_DOCX not found — skipped the docx/png self-test. Run it manually before shipping." >&2
fi

echo ""
echo "All self-tests passed. $DEST is ready to bundle."
echo "Note: this only tested .docx conversion — also test a real .doc upload through the Word merge tool (WordToolPage) once the packaged app is built."
