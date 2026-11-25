
BeforeAll {
    # Pester discovers each test file before it runs it. Load the test harness during
    # the run phase so its setup and helper functions are available to the tests.
    . (Join-Path $PSScriptRoot 'TestSupport.ps1')
    Initialize-WorkTimerTestEnvironment
    Import-WorkTimerFunctions -Name @(
        'Initialize-Storage', 'Write-FileAtomic', 'Get-PropertyValue', 'New-List', 'New-Id',
        'Get-DateKey', 'ConvertFrom-IsoValue', 'ConvertTo-DateKeyText', 'ConvertTo-HistoryRecord'
    )
}

Describe 'Local storage helpers' {
    It 'replaces a file atomically and keeps the prior value when backup is requested' {
    Set-WorkTimerTestStorage -Root $TestDrive
    $path = Join-Path $TestDrive 'state.json'

    Write-FileAtomic -Path $path -Content 'first'
        Write-FileAtomic -Path $path -Content 'second' -KeepBackup

        [IO.File]::ReadAllText($path) | Should -Be 'second'
        [IO.File]::ReadAllText("$path.bak") | Should -Be 'first'
    }

    It 'creates the configured application-data folder when it does not exist' {
        Set-WorkTimerTestStorage -Root $TestDrive

        Initialize-Storage

        Test-Path -LiteralPath $global:AppDirectory -PathType Container | Should -BeTrue
    }
}

Describe 'Saved history conversion' {
    It 'preserves known period data from a saved history record' {
        $source = [pscustomobject]@{
            Id = 'record-1'
            Date = '2026-07-03'
            SessionStart = '2026-07-03T08:00:00.0000000'
            SessionEnd = '2026-07-03T12:00:00.0000000'
            ActiveSeconds = 12600
            ActiveHours = 3.5
            InactiveSeconds = 1800
            InactivePeriods = '10:00-10:30 Manual'
            Reason = 'Ended manually'
            AdjustmentSeconds = 0
            NeedsReview = $false
            Periods = @([pscustomobject]@{
                Type = 'Manual'
                Start = '2026-07-03T10:00:00.0000000'
                End = '2026-07-03T10:30:00.0000000'
            })
        }

        $record = ConvertTo-HistoryRecord $source

        $record.Id | Should -Be 'record-1'
        $record.Date | Should -Be '2026-07-03'
        @($record.Periods).Count | Should -Be 1
        $record.Periods[0].Type | Should -Be 'Manual'
        $record.ActiveSeconds | Should -Be 12600
    }

    It 'retains a null period list for a legacy record whose inactive detail is unknown' {
        $source = [pscustomobject]@{
            Id = 'legacy-1'
            Date = '2026-07-03'
            SessionStart = '2026-07-03T08:00:00.0000000'
            SessionEnd = '2026-07-03T12:00:00.0000000'
            ActiveSeconds = 12600
            ActiveHours = 3.5
            InactiveSeconds = 1800
            InactivePeriods = ''
            Reason = 'Imported legacy log'
            AdjustmentSeconds = 0
            NeedsReview = $false
            Periods = $null
        }

        (ConvertTo-HistoryRecord $source).Periods | Should -BeNullOrEmpty
    }

    It 'uses the saved session start to determine the normalized date when the saved date is ambiguous' {
        $source = [pscustomobject]@{
            Id = 'record-2'
            Date = '10/1/2026'
            SessionStart = '2026-10-01T08:00:00.0000000'
            SessionEnd = '2026-10-01T12:00:00.0000000'
            ActiveSeconds = 14400
            ActiveHours = 4
            InactiveSeconds = 0
            InactivePeriods = ''
            Reason = 'Ended manually'
            AdjustmentSeconds = 0
            NeedsReview = $false
            Periods = @()
        }

        (ConvertTo-HistoryRecord $source).Date | Should -Be '2026-10-01'
    }
}
