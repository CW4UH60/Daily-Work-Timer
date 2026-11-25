# Development Notes

## Application Architecture

Work Timer is intentionally delivered as a single PowerShell application file: `src/WorkTimer.ps1`. The script embeds limited C# interop code for Windows APIs, session notifications, taskbar identity, and shortcut metadata. Keep embedded code and PowerShell changes coordinated.

## Local Data Boundary

Runtime data is stored outside the repository in `%APPDATA%\PowerShellPresenceWorkTimer`. The JSON history file is the source of truth; the Excel workbook is generated from that history. Never use production records as test fixtures or commit them to this repository.

## Validation Scope

`scripts/Test-ScriptSyntax.ps1` uses the PowerShell parser to catch syntax errors without launching Windows Forms. GitHub Actions runs this parser validation on a Windows runner. It does not replace manual testing of UI, presence detection, session events, recovery behavior, file locking, shortcuts, or Excel compatibility.

## Versioning and Releases

No release process is defined yet. Before publishing releases, establish an explicit versioning scheme, release notes process, code-signing approach, distribution method, and license.
