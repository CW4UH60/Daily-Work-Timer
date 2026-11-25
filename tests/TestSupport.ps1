Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:WorkTimerSourcePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'src/WorkTimer.ps1'
$tokens = $null
$script:WorkTimerParseErrors = $null
$script:WorkTimerAst = [System.Management.Automation.Language.Parser]::ParseFile(
    $script:WorkTimerSourcePath,
    [ref]$tokens,
    [ref]$script:WorkTimerParseErrors
)
$script:WorkTimerFunctionDefinitions = @{}
foreach ($definition in $script:WorkTimerAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    $script:WorkTimerFunctionDefinitions[$definition.Name] = $definition.Extent.Text
}

function Initialize-WorkTimerTestEnvironment {
    $global:Culture = [Globalization.CultureInfo]::InvariantCulture
    $global:Utf8 = [Text.UTF8Encoding]::new($false)
    $global:TenthHourSeconds = 360
    $global:PeriodTypes = [ordered]@{
        Manual = 'Manual'
        Idle = 'Idle (no input)'
        Paused = 'Paused'
        Offline = 'Timer not running'
        Asleep = 'Computer asleep'
        Stopped = 'After workday ended'
    }
    $global:State = $null
    $global:History = [Collections.ArrayList]::new()
    $global:ExcelPending = $false
    $global:LastExportAttempt = [DateTime]::MinValue
}

function Import-WorkTimerFunctions {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Name)

    foreach ($functionName in $Name) {
        if (-not $script:WorkTimerFunctionDefinitions.ContainsKey($functionName)) {
            throw "Function '$functionName' was not found in $script:WorkTimerSourcePath."
        }

        # The application is an entry script with startup code. Extracting individual function
        # definitions lets deterministic logic be tested without starting the Windows Forms GUI.
        $definition = $script:WorkTimerFunctionDefinitions[$functionName]
        $pattern = '^\s*function\s+' + [regex]::Escape($functionName) + '\b'
        $regex = [regex]::new($pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Multiline)
        $definition = $regex.Replace($definition, "function global:$functionName", 1)
        $definition = $definition -replace '\$script:', '$global:'
        Invoke-Expression $definition
    }
}

function Get-WorkTimerParseErrors {
    @($script:WorkTimerParseErrors)
}

function Get-WorkTimerFunctionNames {
    @($script:WorkTimerFunctionDefinitions.Keys)
}

function Set-WorkTimerTestStorage {
    param([Parameter(Mandatory)][string]$Root)

    $global:AppDirectory = Join-Path $Root 'data'
    $global:StatePath = Join-Path $global:AppDirectory 'state.json'
    $global:SettingsPath = Join-Path $global:AppDirectory 'settings.json'
    $global:HistoryPath = Join-Path $global:AppDirectory 'work-history.json'
    $global:WorkbookPath = Join-Path $global:AppDirectory 'Work History.xlsx'
    $global:ErrorLogPath = Join-Path $global:AppDirectory 'errors.log'
}

function New-TestPeriod {
    param(
        [Parameter(Mandatory)][DateTime]$Start,
        [Parameter(Mandatory)][DateTime]$End,
        [string]$Type = 'Manual'
    )

    [pscustomobject]@{ Start = $Start; End = $End; Type = $Type }
}

function New-TestState {
    param(
        [Parameter(Mandatory)][DateTime]$Start,
        [object[]]$Periods = @(),
        [AllowNull()][object]$CurrentIdleStart = $null,
        [bool]$Paused = $false,
        [AllowNull()][object]$PauseStart = $null,
        [double]$AdjustmentSeconds = 0
    )

    $periodList = [Collections.ArrayList]::new()
    foreach ($period in $Periods) { [void]$periodList.Add($period) }

    [pscustomobject]@{
        TrackingStarted = $true
        SessionStart = $Start
        Ended = $false
        EndedAt = $null
        Periods = $periodList
        CurrentIdleStart = $CurrentIdleStart
        Paused = $Paused
        PauseStart = $PauseStart
        AdjustmentSeconds = $AdjustmentSeconds
    }
}

function New-TestRecord {
    param(
        [Parameter(Mandatory)][string]$Date,
        [AllowNull()][object]$SessionStart = $null,
        [AllowNull()][object]$SessionEnd = $null,
        [double]$ActiveSeconds = 0,
        [double]$ActiveHours = 0,
        [double]$InactiveSeconds = 0,
        [AllowNull()][object]$Periods = @(),
        [double]$AdjustmentSeconds = 0,
        [string]$Reason = 'Ended manually',
        [bool]$NeedsReview = $false
    )

    [pscustomobject]@{
        Id = [Guid]::NewGuid().ToString('N')
        Date = $Date
        SessionStart = $SessionStart
        SessionEnd = $SessionEnd
        ActiveSeconds = $ActiveSeconds
        ActiveHours = $ActiveHours
        InactiveSeconds = $InactiveSeconds
        InactivePeriods = ''
        Reason = $Reason
        Periods = $Periods
        AdjustmentSeconds = $AdjustmentSeconds
        NeedsReview = $NeedsReview
    }
}
