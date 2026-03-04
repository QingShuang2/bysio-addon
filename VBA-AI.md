# AI-driven VBA add-in build & test

This folder includes a small scaffold for generating and testing VBA add-ins using an automated build script.

Files added
- `src/vba/MathLib.bas` — sample module generated as an example.
- `src/vba/Tests.bas` — simple test runner that writes results to a `TestResults` sheet.
- `tools/build_and_test.ps1` — PowerShell script that imports `*.bas` files, saves `dist/MyAddin.xlam`, runs `RunAllTests`, and writes `dist/test-results.txt`.

How to run (Windows, with Excel installed)
1. Open PowerShell in the repository root.
2. Run:
```powershell
.	ools\build_and_test.ps1
```
3. Inspect `dist\test-results.txt` for test output and `dist\MyAddin.xlam` for the built add-in.

Notes
- Excel must allow programmatic VB project access: Excel → Trust Center → Macro Settings → enable "Trust access to the VBA project object model".
- Running the script requires Excel on the machine (PowerShell COM automation); CI requires a Windows runner with Excel installed.
