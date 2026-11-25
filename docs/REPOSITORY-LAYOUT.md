# Repository Layout

| Path | Purpose |
| --- | --- |
| `src/WorkTimer.ps1` | Application entry point and implementation. |
| `tests/` | Pester tests and a source-function extraction helper. |
| `scripts/Test-ScriptSyntax.ps1` | Parser-based syntax validation for repository PowerShell files. |
| `docs/` | Development, test, and maintenance guidance. |
| `.github/workflows/validate.yml` | Continuous-integration validation workflow. |
| `.github/ISSUE_TEMPLATE/` | Bug-report and feature-request issue forms. |

Generated application data belongs in the per-user `%APPDATA%\PowerShellPresenceWorkTimer` folder and is intentionally excluded from source control.
