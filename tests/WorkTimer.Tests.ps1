#requires -Version 7.2
<#!
.SYNOPSIS
    Repository source-level validation entry point.
.DESCRIPTION
    Kept dependency-free so it can run locally and in GitHub Actions without Pester.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot '..\scripts\Test-ScriptSyntax.ps1')
