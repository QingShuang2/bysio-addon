# PowerPoint VBA Add-in Scaffold

This project builds a PowerPoint add-in (.ppam) from exported VBA modules and runs a small VBA test suite.

## Project structure

- `src/vba/*.bas` - VBA modules imported into a temporary presentation.
- `tools/build_and_test.ps1` - Build and test entry point.
- `tools/lib/powerpoint_build.ps1` - COM helpers for build/import/save.
- `tools/lib/addin_tests.ps1` - Runs `RunAllTests` and reads `GetTestResults`.
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
- `dist/test-results.txt`
