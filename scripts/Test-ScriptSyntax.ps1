#requires -Version 7.2
<#!
.SYNOPSIS
    Parses the Work Timer application source and fails when PowerShell reports syntax errors.
.DESCRIPTION
    This test intentionally does not launch the Windows Forms application. It can run in CI
    and validates only that the source parses as PowerShell.
#>

[CmdletBinding()]
param(
    [string]$Path = (Join-Path $PSScriptRoot '..\src\WorkTimer.ps1')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($resolvedPath, [ref]$tokens, [ref]$errors)

if ($errors.Count -gt 0) {
    $details = $errors | ForEach-Object {
        '{0} (line {1}, column {2})' -f $_.Message, $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber
    }
    throw "PowerShell parser found $($errors.Count) error(s) in '$resolvedPath':`n$($details -join [Environment]::NewLine)"
}

Write-Host "PASS: PowerShell parser found no syntax errors in $resolvedPath"
