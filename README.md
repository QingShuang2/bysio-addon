# Office Add-in (Ribbon commands) — scaffold

This repository contains a minimal Office Add-in scaffold that uses Office.js and Ribbon commands (ShowTaskpane).

Quick start

1. Install dependencies:
   ```bash
   npm install
   ```

2. Start the dev server (serves HTTPS at https://localhost:3000):
   ```bash
   npm start
   ```

3. Sideload the manifest into Excel (web or desktop):
   - Excel for the web: Insert → Add-ins → My Add-ins → "Upload My Add-in" → select `manifest.xml`.
   - Excel Desktop: Insert → Get Add-ins → My Add-ins → "Upload My Add-in" (or use `npx office-addin-debugging`).

Files of interest
- `manifest.xml` — add-in manifest with Ribbon command
- `devserver.js`, `package.json` — local HTTPS dev server (auto-certs)
- `src/taskpane/*`, `src/commands/*`, `src/assets/*` — web assets
- `tools/build_and_test.ps1` — (Windows) imports `src/vba/*.bas`, builds `dist/MyAddin.xlam`, runs `RunAllTests`, writes `dist/test-results.txt`.

Notes
- The dev server generates a self-signed certificate at runtime; for Desktop Excel you may need to trust it.
- `tools/build_and_test.ps1` automates building a VBA add-in by importing exported `.bas` modules into Excel. This operation requires that Excel allow programmatic access to the VBA project object model.

Registry change applied while running the build script
- To allow the script to import `.bas` modules the per-user registry key `AccessVBOM` was set to `1` under any found Office Excel security key (for example `HKCU:\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM`).
- If you prefer to revert this change, run the following PowerShell snippet (reverts the key to `0` for all detected Office versions):

```powershell
Get-ChildItem HKCU:\Software\Microsoft\Office | ForEach-Object {
  $sec = Join-Path $_.PSPath 'Excel\Security'
  if (Test-Path $sec) {
    Set-ItemProperty -Path $sec -Name 'AccessVBOM' -Value 0 -Type DWord
    Write-Host "Reverted AccessVBOM in $sec"
  }
}
```

Use the build script (Windows with Excel installed)
```powershell
.
\tools\build_and_test.ps1
```

The script outputs `dist\MyAddin.xlam` and `dist\test-results.txt`.
