# Stages a pruned copy of the local LibreOffice install into resources/libreoffice-win/,
# for electron-builder's win.extraResources to bundle into the packaged app —
# so an office clerk's Windows machine never needs its own separate
# LibreOffice install for accurate document preview/print/PDF export.
# Run this once before `npm run dist`, or whenever LibreOffice is upgraded on
# this machine. resources/libreoffice-win/ is gitignored, never committed.
#
# The exclusion list below was derived empirically on macOS (removing
# candidates, then running real conversions, one category at a time — not
# guessed from file names) and translated to the equivalent Windows DLL
# names. Two categories LOOKED safe to remove but turned out to be load-time
# hard dependencies, discovered only by actually testing:
#   - The legacy-format import filters (mwaw/staroffice/wpg/wpd/etonyek/wps —
#     AppleWorks/Keynote/StarOffice/WordPerfect/MS Works) are statically
#     linked INTO the core merged library itself. Removing them crashes
#     soffice at startup even though this app never opens those formats.
#   - "Draw" (sd*) is not just for .odg files — it's the component
#     LibreOffice's own PDF->PNG rasterization goes through internally
#     (core/docxToPdf.ts's docxToPageImages, the app's entire accurate
#     preview/print pipeline). Removing it broke PDF->image export outright.
# Removed instead (verified NOT to be load-time dependencies, on macOS):
# Calc, Math, VBA object support, the PostgreSQL Base connector — nothing
# this app's Writer-only docx/doc -> pdf/png pipeline ever touches.
#
# IMPORTANT — this script's exclusion list has only been verified on macOS.
# Windows DLL naming/dependency resolution can differ, so the SelfTest step
# below actually exercises the staged copy after pruning (docx convert, doc
# convert — the WordToolPage "Word merge" tool accepts legacy .doc uploads —
# and PDF->PNG, matching the app's real usage) and fails loudly if anything's
# missing, rather than silently shipping a broken bundle. Don't skip it.

param(
  [string]$Source = "C:\Program Files\LibreOffice"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$Dest = Join-Path $RepoRoot "resources\libreoffice-win"

if (-not (Test-Path $Source)) {
  Write-Error "LibreOffice not found at $Source — install it, or pass -Source <path>."
  exit 1
}

Write-Host "Staging a pruned copy of $Source -> $Dest ..."
if (Test-Path $Dest) { Remove-Item -Recurse -Force $Dest }
New-Item -ItemType Directory -Path $Dest | Out-Null

# GUI-only / interactive-only asset folders — never touched by --headless --convert-to.
$ExcludeDirs = @('help', 'extensions', 'gallery', 'template', 'wizards', 'firebird', 'java')
robocopy $Source $Dest /E /XD $ExcludeDirs /XF "images_*.zip" "CREDITS.fodt" "*.ico" /NFL /NDL /NJH /NJS | Out-Null

# Toolbar icon-theme archives — headless mode never renders a toolbar.
Get-ChildItem -Path $Dest -Recurse -Filter "images_*.zip" -ErrorAction SilentlyContinue | Remove-Item -Force

# Calc / Math / VBA-object / PostgreSQL-Base-connector DLLs — empirically not
# load-time dependencies of the core engine (unlike the legacy filters and
# Draw, kept above). Wildcards match both the base module and its UI/filter
# siblings (e.g. sc.dll, scui.dll, scfilt.dll, scd.dll, scn.dll) without
# touching sd*/sw* (Draw/Writer, both required).
$ProgramDir = Join-Path $Dest "program"
if (Test-Path $ProgramDir) {
  $patterns = @('sc.dll', 'scui.dll', 'scfilt.dll', 'scd.dll', 'scn.dll', 'sm.dll', 'smd.dll', 'vbaobj.dll', 'vbaswobj.dll', 'postgresql-sdbc*.dll', 'mysqlc*.dll', 'mysql.dll')
  foreach ($pattern in $patterns) {
    Get-ChildItem -Path $ProgramDir -Filter $pattern -ErrorAction SilentlyContinue | Remove-Item -Force
  }
}

$sizeMB = [math]::Round((Get-ChildItem -Path $Dest -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB)
Write-Host "Staged size: $sizeMB MB"

# --- Self-test: prove the pruned copy actually still works before trusting it. ---
$soffice = Join-Path $ProgramDir "soffice.exe"
if (-not (Test-Path $soffice)) {
  Write-Error "SELF-TEST FAILED: soffice.exe missing from the staged copy at $soffice."
  exit 1
}

$TestDir = Join-Path $env:TEMP "libreoffice-win-selftest"
if (Test-Path $TestDir) { Remove-Item -Recurse -Force $TestDir }
New-Item -ItemType Directory -Path $TestDir | Out-Null

function Test-Conversion($label, $inputFile, $convertTo) {
  Write-Host "Self-test: $label ..."
  & $soffice --headless --convert-to $convertTo --outdir $TestDir $inputFile 2>&1 | Out-Null
  $outName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
  $outExt = ($convertTo -split ':')[0]
  $outFile = Join-Path $TestDir "$outName.$outExt"
  if (-not (Test-Path $outFile) -or (Get-Item $outFile).Length -eq 0) {
    Write-Error "SELF-TEST FAILED: $label produced no output. The pruned bundle at $Dest is broken — do not ship it. Check which DLL is missing by running soffice.exe directly from a terminal and reading its error."
    exit 1
  }
  Write-Host "  OK ($((Get-Item $outFile).Length) bytes)"
  return $outFile
}

$sampleDocx = Join-Path $RepoRoot "resources\work-order-template.docx"
if (Test-Path $sampleDocx) {
  $pdf = Test-Conversion "docx -> pdf" $sampleDocx "pdf"
  Test-Conversion "pdf -> png (the app's accurate preview/print path)" $pdf 'png:draw_png_Export:{"PixelWidth":{"type":"long","value":"2480"},"PixelHeight":{"type":"long","value":"3508"}}' | Out-Null
} else {
  Write-Warning "resources\work-order-template.docx not found — skipped the docx/png self-test. Run it manually before shipping."
}

Write-Host ""
Write-Host "All self-tests passed. $Dest is ready to bundle." -ForegroundColor Green
Write-Host "Note: this only tested .docx conversion — also test a real .doc upload through the Word merge tool (WordToolPage) once the packaged app is built, since that's the other format this bundle needs to keep working." -ForegroundColor Yellow
