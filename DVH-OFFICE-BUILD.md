# DVH Office - light interface, Gemini AI and DVH functions

Based on genspark-ai/genoffice at d9cd4895. Upstream licenses and notices are retained.

## Current build

- Installer: `apps/shell/release/dvh-gemini/DVH Office Setup 0.10.0.exe`
- Portable application: `apps/shell/release/dvh-gemini/win-unpacked/DVH Office.exe`
- Keep the whole `win-unpacked` directory together.
- Older builds under `release`, `release/no-genspark`, `release/ui-clean`, `release/no-ai-light`, `release/font-fix`, and `release/dvh-home` predate Gemini AI. Use `dvh-gemini`.
- Save documents and close the installed DVH Office before running the new installer. This build has not automatically replaced the running installation.

## Scope

AI chat and Review → Translate are enabled with Gemini. Settings → AI Model selects a Gemini model and lets the user test or replace the API key. On Windows the app automatically reads the Gemini key already saved by DVH-Tool in `HKCU\Software\DVH_Tool\ApiKeys`; a newly entered key is encrypted for the current Windows user in DVH Office's user data. The key is not packaged or written to `ai-settings.json`. The AI chat backend accepts Gemini only and does not offer Genspark sign-in. Search and media generation retain their separate configuration and are not covered by the Gemini chat key.

Sheets Home now includes DVH Excel's center-across-selection, justify, shrink-to-fit, merged-cell row fitting, custom text indent, and border preset commands. The Settings dialog edits indent and border presets. Initial defaults were copied from the three indent and two border presets saved locally by DVH-Tool on this machine. User edits are stored in DVH Office's local settings; the original Excel add-in and its registry values are unchanged. Formula cells are skipped by custom text indent.

`packages/ui/src/product-features.ts` records the product policy. Genspark execution and credential access remain removed from the earlier work.

All six editors and the shell use the light theme. Existing dark/system settings cannot switch the desktop back to dark. The native window theme is light, and the Word dark-page control is disabled. Document-authored colors are preserved.

## Verification

- Built all editors and the Windows x64 shell/installer.
- TypeScript checked the workspaces; affected editors/shell checked again after final corrections.
- Lint: no errors; 12 existing warnings.
- Packaged Gemini build: Settings → AI Model shows only Gemini; the existing DVH-Tool key is imported; a live one-shot chat request succeeded once (a later repeat returned a temporary busy response); saving stores an encrypted `gemini-key.bin` and leaves `ai-settings.json` without a plaintext key. The packaged Sheets Insert Function dialog lists 39 DVH functions. `DVH.Evaluate("2*3*4")`, `DVH.LastRow(A:A)` and `DVH.LastCol(A7)` displayed and saved their expected cached results (`2*3*4=24`, `12`, `2`) in XLSX. Driver: `scripts/drivers/driver.dvh-gemini.mjs`.
- The Gemini package was opened in all six editors (Sheets, Docs, Slides, PDF, Markdown, HTML). Each remained light even with a saved dark preference and dark OS color scheme, and none showed Genspark. Driver and report: `scripts/drivers/driver.dvh-gemini-all-editors.mjs`, `.build-logs/dvh-gemini-all-editors.json`.
- Word opens and renders a DOCX; Sheets opens the XLSX.
- Earlier packaged CLI checks created/read DOCX and XLSX, including 10 + 20 = 30.
- Packaged Sheets UI checks applied all six DVH Home formatting commands to an XLSX, verified saved alignment, shrink and border styles, checked a formula cell survived indent and border operations, and tested one-step Undo. The presets dialog displayed three indent and two border templates. Logs and screenshots: `.build-logs/dvh-home-ui.log`, `.build-logs/dvh-home-ribbon.png`, `.build-logs/dvh-home-settings.png`.
- Screenshots and UI test report: `.build-logs/no-ai-light-*.png` and `.build-logs/no-ai-light-ui.json`.

This remains an unsigned local trial. Installer execution and comprehensive Office compatibility were not tested. The full upstream suite is not green on this machine: previous broad runs encountered Windows path/permission/symlink/default-app assumptions and a PDF-conversion timeout. Focused product checks are recorded in `.build-logs`.

## Rebuild

Run `./Build-DVH-Office.ps1`. It uses the existing portable Node, Rust GNU and MinGW toolchains under `.local-tools`, plus the compiled Windows OCR helper and installed npm dependencies. Toolchains, logs and generated binaries are not committed. These prerequisites must remain on this machine.

## Font menu freeze fix (2026-09-28)

The Docs font picker rendered every installed font name in its own font. On this machine, loading 2,213 menu entries stalled a renderer query for about 26.7 seconds. The picker now uses the UI font for its labels; choosing a font still sets the document text style.

Packaged-app before/after tests used the same small DOCX, fresh profiles and a 10-second pause after opening the picker. The subsequent menu query took about 26.7 seconds before and 13 milliseconds after. Arial and Times New Roman were selected successfully after the change. Docs remained light with no AI controls. Logs: `.build-logs/font-before-timing.log` and `.build-logs/font-after-timing.log`; driver: `scripts/drivers/driver.dvh-font-fixed.mjs`.

Docs typecheck and targeted Ribbon ESLint passed. The Sheets picker also loaded 2,202 entries and selected Arial successfully during diagnosis; no Sheets font changes were made. This verifies the reproduced Docs menu stall, not all possible workbook/font issues.
