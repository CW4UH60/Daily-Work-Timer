# Testing Guide

## Automated coverage

The Pester suite tests deterministic, non-UI behavior directly from `src/WorkTimer.ps1` without launching the Windows Forms application:

- PowerShell parser validity and required calculation/export function presence.
- Interval clipping and union logic, including overlapping inactive periods.
- Active-work totals, current idle time, pauses, and manual adjustments.
- Rounded hours, next-tenth-hour countdown boundaries, and inactive-period summaries.
- ISO timestamp and date-key conversion, including legacy-record behavior.
- Atomic-file replacement/backup behavior.
- Generated Excel workbook ZIP structure, table metadata, formulas, headers, and saved record content.

Run it with:

```powershell
pwsh -NoProfile -File .\scripts\Test-ScriptSyntax.ps1
Invoke-Pester -Path .\tests -Output Detailed
```

## Manual Windows test checklist

Perform applicable checks in an isolated test account or with disposable test data before release:

### Startup and UI

- Start the application in an interactive Windows desktop session.
- Confirm it starts in an STA PowerShell process and opens one window without an unwanted console window when launched through an app-created shortcut.
- Confirm a second launch activates or reports the existing instance rather than creating a competing writer.
- Check display scaling, keyboard navigation, dialogs, and close confirmation.

### Presence and workday behavior

- Verify tracking starts after the first keyboard/mouse input and after sign-in/unlock.
- Set a short idle threshold; verify idle begins after the threshold and ends on user input.
- Pause and resume; verify the paused interval is excluded once.
- Add, edit, overlap, and remove manual inactive periods; confirm totals update correctly.
- End and reopen a workday; confirm the post-end interval is represented as inactive by default.
- Leave a workday open across midnight; verify the review workflow flags the prior day.

### Recovery and storage

- Close and reopen during an active day; test both **Count as Work** and **Exclude** recovery choices.
- Sleep/hibernate and resume; verify the detected gap is excluded and remains editable.
- Verify settings, state, history, backup, and error handling using disposable copies of the application-data folder.
- Verify legacy CSV import only with non-production test files.

### Export and shortcuts

- Create history records and open the generated `Work History.xlsx` in Microsoft Excel.
- Confirm column values, day and hours-worked formulas, filters, totals, and expected sort order.
- Keep the workbook open while updating history; confirm the app indicates a pending export and updates after the file is closed.
- Test Start Menu and Start-with-Windows shortcut operations in a non-production user profile.

Record Windows version, PowerShell version, test date, scenarios exercised, and results for each release candidate.
