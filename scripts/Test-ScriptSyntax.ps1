#requires -Version 7.2

<#
.SYNOPSIS
    Parses PowerShell source and test files and fails when syntax errors are found.

.DESCRIPTION
    This validation does not start the Windows Forms application. It checks all
    PowerShell files in src, tests, and scripts so it is suitable for CI.
#>

[CmdletBinding()]
param(
    [string[]]$Path = @(
        'src/*.ps1',
        'tests/*.ps1',
        'scripts/*.ps1'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$seenPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($item in $Path) {
    $candidate = if ([System.IO.Path]::IsPathRooted($item)) {
        $item
    }
    else {
        Join-Path $repositoryRoot $item
    }

    $resolved = @(
        Get-ChildItem -Path $candidate -File -ErrorAction SilentlyContinue
    )

    if ($resolved.Count -eq 0) {
        throw "No PowerShell files matched '$item'."
    }

    foreach ($file in $resolved) {
        if ($seenPaths.Add($file.FullName)) {
            [void]$files.Add($file)
        }
    }
}

$failures = [System.Collections.Generic.List[string]]::new()

foreach ($file in $files) {
    $tokens = $null
    $errors = $null

    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName,
        [ref]$tokens,
        [ref]$errors
    )

    foreach ($error in $errors) {
        [void]$failures.Add(('{0}:{1}:{2} {3}' -f $file.FullName, $error.Extent.StartLineNumber, $error.Extent.StartColumnNumber, $error.Message))
    }

    # A lone backslash is syntactically valid PowerShell but attempts to invoke a
    # command named "\\" at runtime. Treat it as a source error to catch accidental
    # leading backslashes in generated files.
    foreach ($match in @(Select-String -LiteralPath $file.FullName -Pattern '^\s*\\\s*$' -AllMatches)) {
        [void]$failures.Add(('{0}:{1}:1 Unexpected standalone backslash command.' -f $file.FullName, $match.LineNumber))
    }

    if ($errors.Count -eq 0) {
        Write-Host "PASS  $($file.FullName)"
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "PowerShell syntax validation failed for $($failures.Count) error(s)."
}

Write-Host "Validated $($files.Count) PowerShell file(s)."
