# PowerPoint VBA Add-in (Minimal Ribbon)

This project builds a PowerPoint add-in (.ppam) with a custom ribbon tab and a single button.

## Project structure

- `src/vba/Ribbon.bas` - Ribbon callback for one button.
- `tools/build_and_test.ps1` - Build entry point.
- `tools/lib/powerpoint_build.ps1` - COM helpers for build/import/save.
- `tools/lib/ribbon_openxml.ps1` - Injects ribbon XML into the .ppam package.
- `dist/` - Output folder for built add-in and test report.

## Requirements

- Windows with Microsoft PowerPoint installed.
- Trust Center setting enabled:
  - PowerPoint -> Trust Center -> Macro Settings -> Trust access to the VBA project object model.

## Run

```powershell
.\tools\build_and_test.ps1
```

Outputs:

- `dist/MyPowerPointAddIn.ppam`

When loaded in PowerPoint, the add-in adds a `Bysio` tab with a `1 + 1` button.
Clicking it shows: `1 + 1 = 2`.
