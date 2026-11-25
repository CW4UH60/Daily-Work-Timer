# Contributing to Work Timer

Thank you for improving Work Timer. This repository contains a Windows-only, local-first PowerShell desktop application.

## Before You Start

- Discuss significant feature changes in an issue before investing substantial implementation effort.
- Keep the application compatible with **PowerShell 7.2+** and Windows desktop sessions.
- Do not add telemetry, network calls, external dependencies, or administrator-rights requirements without explicit discussion and documentation.
- Do not commit user work-history files, logs, generated workbooks, or other local data.

## Development Setup

1. Install PowerShell 7.2 or later on Windows.
2. Clone the repository.
3. Launch the application from the repository root:

   ```powershell
   pwsh -NoProfile -STA -File .\src\WorkTimer.ps1
   ```

4. Run source-level validation before submitting a pull request:

   ```powershell
   pwsh -NoProfile -File .\scripts\Test-ScriptSyntax.ps1
   ```

The parser test does not open the application or validate Windows Forms behavior. Manually exercise affected workflows on Windows.

## Change Expectations

For user-visible changes, update `README.md` and add a concise entry under **Unreleased** in `CHANGELOG.md`. Consider the following manual checks when relevant:

- First-input start and active-time calculation.
- Pause/resume and manual inactive periods.
- Idle detection and configurable threshold.
- Sleep, restart, and closed-app recovery prompts.
- Workday completion, rollover, and history correction.
- JSON persistence, backups, and Excel workbook export.
- Start menu/startup shortcut behavior and single-instance protection.

## Pull Requests

- Keep changes focused and explain the user impact.
- Describe the validation you performed and its results.
- Do not include sensitive or personally identifying information in issues, commits, test fixtures, screenshots, or logs.
- Ensure GitHub Actions completes successfully.
