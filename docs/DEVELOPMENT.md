# Development Guide

## Prerequisites

- Windows
- PowerShell 7.2 or later
- Git
- Pester 5.x for the automated test suite

## Run the application

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\src\WorkTimer.ps1
```

The application requires Windows Forms and Windows input/session APIs. Run it in an interactive Windows desktop session.

## Run validation

```powershell
pwsh -NoProfile -File .\scripts\Test-ScriptSyntax.ps1
Invoke-Pester -Path .\tests -Output Detailed
```

The parser script validates every `.ps1` file under `src`, `tests`, and `scripts` by default. The Pester suite loads selected non-UI function definitions from the application source without starting the GUI.

## Design boundaries

- `src/WorkTimer.ps1` is the application entry point and currently contains both UI and business logic.
- Tests focus on deterministic business logic and file-package output. They do not automate native Windows Forms controls or real user input/session transitions.
- Do not use real `%APPDATA%\PowerShellPresenceWorkTimer` data for automated tests. Tests use Pester's temporary test drive.
- The Excel export is generated as an Office Open XML ZIP package. Test both its package contents and, when changing export behavior, its opening in Microsoft Excel or another compatible spreadsheet application.

See [TESTING.md](TESTING.md) for the required manual validation areas.
