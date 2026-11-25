# Work Timer

A local, presence-aware desktop work timer for Windows. **Work Timer** starts tracking on your first keyboard or mouse activity, calculates active time by subtracting inactive intervals from the workday, and maintains an editable daily history with an Excel export.

> **Privacy note:** Idle detection uses Windows' last-input timestamp to determine elapsed inactivity. The script does not capture keystrokes, mouse content, screenshots, or send timekeeping data to a remote service.

## Features

- **Automatic start:** Begins a workday when keyboard or mouse input is first detected.
- **Presence-aware tracking:** Detects inactivity after a configurable no-input threshold.
- **Accurate active-time calculation:** Active time is the workday elapsed time minus the union of inactive periods, preventing overlap from being counted twice.
- **Editable inactive periods:** Review, add, modify, or remove manual, idle, paused, offline, sleep, and post-workday intervals.
- **Pause and end controls:** Pause/resume the timer or explicitly end the workday.
- **Sleep and restart recovery:** Detects sleep/hibernate gaps and asks whether a period while the app was closed should count as work or be excluded.
- **Daily history:** Browse, filter, and edit prior workdays; incomplete days can be reviewed and corrected.
- **Excel work log:** Generates an `Work History.xlsx` workbook with a table, totals, and formulas—without requiring Microsoft Excel or third-party PowerShell modules to create it.
- **Single-instance protection:** Ensures only one Work Timer process runs for each Windows user.
- **Convenient shortcuts:** Create a Start menu shortcut or start automatically at sign-in without administrator rights.
- **Native desktop experience:** Windows Forms UI, app icon, taskbar identity, and high-DPI support.

## Requirements

- Windows
- [PowerShell 7.2](https://learn.microsoft.com/powershell/)
- A desktop session capable of running Windows Forms
- No administrator rights required
- No external PowerShell modules required
- Microsoft Excel is **optional**: the script generates the `.xlsx` file directly; Excel or another spreadsheet application is only needed to view it.

## Installation

1. Download or clone this repository.
2. Save the script as `WorkTimer.ps1` in a location that will remain available. If your copy is named `WorkTimer_v2.txt`, rename it to `WorkTimer.ps1`.
3. Open PowerShell 7 and run:

   ```powershell
   cd C:\Path\To\WorkTimer
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\WorkTimer.ps1
   ```

   The `-ExecutionPolicy Bypass` value above applies only to that launch. Alternatively, use your organization's approved PowerShell execution-policy and code-signing process.

4. Work Timer opens as a desktop application. It waits for your first keyboard or mouse activity before starting the workday.

### Start without a console window

The app-created Start menu and startup shortcuts launch Work Timer without showing a PowerShell console. To launch it manually in the same manner, use:

```powershell
conhost.exe --headless pwsh.exe -NoProfile -STA -WindowStyle Hidden -File "C:\Path\To\WorkTimer\WorkTimer.ps1"
```

On supported Windows versions, the application uses `conhost --headless`; otherwise, its shortcut logic falls back to launching `pwsh` directly.

## How It Works

### Active-time calculation

```text
Active work = time since workday start − union of all inactive periods
```

Inactive periods may be created automatically or manually:

| Type | When it is used |
| --- | --- |
| **Idle (no input)** | No keyboard or mouse input exceeds the configured threshold. |
| **Paused** | You pause the timer. |
| **Manual** | You add a period in the Idle Periods editor. |
| **Timer not running** | You choose to exclude a gap after reopening the app. |
| **Computer asleep** | A sleep or hibernation gap is detected. |
| **After workday ended** | You reopen an already ended workday; the intervening time is excluded by default. |

All periods remain editable, and totals are recalculated after each change.

### Main window

The main window provides four primary actions:

- **End Workday** — completes the current day and saves its history record.
- **Pause Timer** — pauses or resumes time tracking.
- **Idle Periods** — reviews and edits inactive intervals for the current day.
- **Work History** — filters, reviews, corrects, and opens past records in Excel.

### Advanced tools

Press **Ctrl+Shift+A** while the main window is focused to open Advanced Tools. Available actions include:

- Edit the current workday's start time.
- Reopen a completed workday.
- Reset today's in-progress record.
- Change the idle threshold from 1 to 1,440 minutes.
- Enable or disable start-at-sign-in.
- Create a Start menu shortcut.
- Open the data folder, rebuild the Excel log, or view the error log.

## Data Storage and Backup

Work Timer stores data under:

```text
%APPDATA%\PowerShellPresenceWorkTimer
```

| File | Purpose |
| --- | --- |
| `state.json` | Current in-progress workday. |
| `settings.json` | User settings, including idle threshold. |
| `work-history.json` | Primary saved work-history data. |
| `work-history.json.bak` | Backup retained during atomic history writes. |
| `Work History.xlsx` | Generated spreadsheet log. |
| `errors.log` | Diagnostic error log. |

The JSON history file is the source of truth. The Excel workbook is regenerated whenever history changes. If `Work History.xlsx` is open and locked, Work Timer retains the update and retries after the file becomes available.

### Backing up or moving data

To preserve history before changing computers or reinstalling Windows, close Work Timer and copy the entire data folder above to a secure backup location. Restore it to the same `%APPDATA%\PowerShellPresenceWorkTimer` location for the applicable user.

## Usage Notes

- **Closing the app does not track time.** If you close Work Timer during an active workday, it asks for confirmation. When reopened, it prompts you to count the gap as work or exclude it.
- **Sleep/hibernate is excluded automatically.** The detected interval can be adjusted later in **Idle Periods**.
- **End the workday when possible.** If a day rolls over without being ended, Work Timer flags the prior day for review so you can enter the actual end time or total hours worked.
- **The default idle threshold is 480 minutes.** Adjust it in **Ctrl+Shift+A → Advanced Tools**.
- **This is a personal local timekeeping utility.** Verify totals against your organization's timekeeping, attendance, and records requirements before using them for official reporting.

## Troubleshooting

### The script says it requires Windows

Work Timer depends on Windows Forms and Windows APIs for last-input and session-change detection. Run it on Windows only.

### PowerShell says scripts are disabled

Launch it with the process-scoped command shown in [Installation](#installation), or follow your organization's approved execution-policy and signing procedures. Do not weaken system-wide execution policy solely to run this script.

### A second copy will not open

Only one instance runs per Windows user to prevent conflicting writes to the same state files. Select the existing Work Timer window from the taskbar or close it before launching another instance.

### The Excel workbook is not updating

Close `Work History.xlsx` in Excel or any application that has locked it. Work Timer retries pending exports automatically; you can also use **Ctrl+Shift+A → Rebuild Excel Log**.

### I need to correct a prior day

Open **Work History**, select the record, and choose **Edit Day**. For a day that was not ended before rollover, follow the prompted review workflow.

### I need diagnostic details

Open **Ctrl+Shift+A → View Error Log**, or review:

```text
%APPDATA%\PowerShellPresenceWorkTimer\errors.log
```

## Project Notes

- The application is written as a single PowerShell script and includes small embedded C# interop helpers for Windows APIs, session notifications, taskbar identity, and shortcut behavior.
- The workbook export is created as an Office Open XML package; it does not automate the Excel application.
- The script can import a legacy CSV work log if one is present in its data folder.

## Contributing

Contributions are welcome. Before submitting a change:

1. Keep the application compatible with PowerShell 7.2 and Windows.
2. Preserve local-first behavior and avoid adding telemetry or external dependencies without explicit documentation.
3. Test start/stop, pause, idle detection, sleep/restart recovery, history editing, and Excel export.
4. Update this README when user-visible behavior, stored data, or requirements change.
