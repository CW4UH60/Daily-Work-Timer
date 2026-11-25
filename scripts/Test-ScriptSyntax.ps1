[CmdletBinding()]
param(
    [string[]]$Path = @('src/*.ps1', 'tests/*.ps1', 'scripts/*.ps1')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

foreach ($item in $Path) {
    $candidate = if ([System.IO.Path]::IsPathRooted($item)) { $item } else { Join-Path $repositoryRoot $item }
    $resolved = @(Get-ChildItem -Path $candidate -File -ErrorAction SilentlyContinue)
    if ($resolved.Count -eq 0) {
        throw "No PowerShell files matched '$item'."
    }
    foreach ($file in $resolved) {
        if (-not $files.FullName.Contains($file.FullName)) { [void]$files.Add($file) }
    }
}

$failures = [System.Collections.Generic.List[string]]::new()
foreach ($file in $files) {
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
    if ($errors.Count -eq 0) {
        Write-Host "PASS  $($file.FullName)"
        continue
    }

    foreach ($error in $errors) {
        [void]$failures.Add(('{0}:{1}:{2} {3}' -f $file.FullName, $error.Extent.StartLineNumber, $error.Extent.StartColumnNumber, $error.Message))
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "PowerShell syntax validation failed for $($failures.Count) error(s)."
}

Write-Host "Validated $($files.Count) PowerShell file(s)."
