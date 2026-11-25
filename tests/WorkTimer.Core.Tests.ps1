
BeforeAll {
    # Pester discovers each test file before it runs it. Load the test harness during
    # the run phase so its setup and helper functions are available to the tests.
    . (Join-Path $PSScriptRoot 'TestSupport.ps1')
    Initialize-WorkTimerTestEnvironment
    Import-WorkTimerFunctions -Name @(
        'Get-DateKey', 'Format-Time', 'Format-Duration', 'Get-RoundedHours', 'Get-SecondsToNextTenth',
        'Measure-UnionSeconds', 'Get-WorkSummary', 'Get-PeriodSummaryText', 'ConvertTo-IsoText',
        'ConvertFrom-IsoValue', 'ConvertTo-DateKeyText', 'Get-RecordInactiveLump',
        'Measure-RecordActiveSeconds', 'Get-EndTimeForActiveSeconds'
    )
}

Describe 'Work Timer source contract' {
    It 'parses without PowerShell parser errors' {
        @(Get-WorkTimerParseErrors).Count | Should -Be 0
    }

    It 'contains the functions that implement core calculation and export behavior' {
        $names = Get-WorkTimerFunctionNames
        foreach ($required in @('Measure-UnionSeconds', 'Get-WorkSummary', 'Measure-RecordActiveSeconds', 'Export-HistoryWorkbook')) {
            $names | Should -Contain $required
        }
    }
}

Describe 'Inactive interval and active-time calculations' {
    It 'clips intervals to the workday and counts overlaps once' {
        $day = [datetime]'2026-07-01T00:00:00'
        $intervals = @(
            (New-TestPeriod $day.AddHours(7) $day.AddHours(8.25)),
            (New-TestPeriod $day.AddHours(8.5) $day.AddHours(9.5)),
            (New-TestPeriod $day.AddHours(9) $day.AddHours(10)),
            (New-TestPeriod $day.AddHours(11) $day.AddHours(12.5))
        )

        $seconds = Measure-UnionSeconds -Intervals $intervals -RangeStart $day.AddHours(8) -RangeEnd $day.AddHours(12)

        $seconds | Should -Be 9900
    }

    It 'ignores zero-length and out-of-range intervals' {
        $day = [datetime]'2026-07-01T00:00:00'
        $intervals = @(
            (New-TestPeriod $day.AddHours(7) $day.AddHours(8)),
            (New-TestPeriod $day.AddHours(9) $day.AddHours(9)),
            (New-TestPeriod $day.AddHours(13) $day.AddHours(14))
        )

        (Measure-UnionSeconds -Intervals $intervals -RangeStart $day.AddHours(9) -RangeEnd $day.AddHours(12)) | Should -Be 0
    }

    It 'includes current idle and pause intervals without double-counting their overlap' {
        $day = [datetime]'2026-07-01T00:00:00'
        $state = New-TestState -Start $day.AddHours(8) `
            -Periods @((New-TestPeriod $day.AddHours(9) $day.AddHours(10) 'Manual')) `
            -CurrentIdleStart $day.AddHours(9.5) `
            -Paused $true -PauseStart $day.AddHours(10.75)

        $summary = Get-WorkSummary -State $state -At $day.AddHours(12)

        $summary.InactiveSeconds | Should -Be 10800
        $summary.ActiveSeconds | Should -Be 3600
    }

    It 'applies a migrated adjustment only after inactive time is removed' {
        $day = [datetime]'2026-07-01T00:00:00'
        $state = New-TestState -Start $day.AddHours(8) `
            -Periods @((New-TestPeriod $day.AddHours(9) $day.AddHours(10) 'Manual')) `
            -AdjustmentSeconds 300

        (Get-WorkSummary -State $state -At $day.AddHours(11)).ActiveSeconds | Should -Be 7500
    }

    It 'never returns negative active or inactive time when the end precedes the start' {
        $day = [datetime]'2026-07-01T00:00:00'
        $state = New-TestState -Start $day.AddHours(9)

        $summary = Get-WorkSummary -State $state -At $day.AddHours(8)

        $summary.ActiveSeconds | Should -Be 0
        $summary.InactiveSeconds | Should -Be 0
    }
}

Describe 'Work-time formatting and editing helpers' {
    It 'rounds hours at the intended half-tenth boundary' {
        Get-RoundedHours 179 | Should -Be 0
        Get-RoundedHours 180 | Should -Be 0.1
        Get-RoundedHours 540 | Should -Be 0.2
    }

    It 'calculates the countdown to the next displayed tenth hour' {
        Get-SecondsToNextTenth 0 | Should -Be 180
        Get-SecondsToNextTenth 180 | Should -Be 360
        Get-SecondsToNextTenth 539 | Should -Be 1
    }

    It 'formats durations for sub-minute, minute, and hour ranges' {
        Format-Duration 59 | Should -Be '0 min'
        Format-Duration 3599 | Should -Be '59 min'
        Format-Duration 3660 | Should -Be '1 h 01 min'
    }

    It 'summarizes only periods that overlap the selected range' {
        $day = [datetime]'2026-07-01T00:00:00'
        $periods = @(
            (New-TestPeriod $day.AddHours(7.5) $day.AddHours(8.25) 'Manual'),
            (New-TestPeriod $day.AddHours(9) $day.AddHours(9.5) 'Paused'),
            (New-TestPeriod $day.AddHours(12) $day.AddHours(13) 'Idle')
        )

        Get-PeriodSummaryText -Periods $periods -RangeStart $day.AddHours(8) -RangeEnd $day.AddHours(12) |
            Should -Be '08:00-08:15 Manual; 09:00-09:30 Paused'
    }
}

Describe 'History record calculations' {
    It 'uses a fixed inactive lump for a legacy record with no period list' {
        $day = [datetime]'2026-07-01T00:00:00'
        $record = New-TestRecord -Date '2026-07-01' -SessionStart $day.AddHours(8) -SessionEnd $day.AddHours(17) `
            -ActiveSeconds 25200 -ActiveHours 7 -InactiveSeconds 7200 -Periods $null

        (Measure-RecordActiveSeconds -Record $record -Start $day.AddHours(8) -End $day.AddHours(12)) | Should -Be 7200
        (Get-EndTimeForActiveSeconds -Record $record -Start $day.AddHours(8) -TargetSeconds 10800) | Should -Be $day.AddHours(13)
    }

    It 'finds an end time that skips known inactive periods' {
        $day = [datetime]'2026-07-01T00:00:00'
        $record = New-TestRecord -Date '2026-07-01' -Periods @((New-TestPeriod $day.AddHours(10) $day.AddHours(11) 'Paused'))

        (Get-EndTimeForActiveSeconds -Record $record -Start $day.AddHours(8) -TargetSeconds 10800) | Should -Be $day.AddHours(12)
    }
}

Describe 'Date and timestamp conversion' {
    It 'round-trips an unspecified local timestamp in ISO form' {
        $value = [datetime]'2026-07-01T08:15:30.1234567'

        (ConvertFrom-IsoValue (ConvertTo-IsoText $value)) | Should -Be $value
    }

    It 'normalizes recognized date formats and uses the start date as an ambiguity fallback' {
        ConvertTo-DateKeyText -Value '2026/7/3' -Fallback $null | Should -Be '2026-07-03'
        ConvertTo-DateKeyText -Value '10/1/2026' -Fallback ([datetime]'2026-10-01T08:00:00') | Should -Be '2026-10-01'
    }

    It 'treats null and blank serialized timestamps as missing values' {
        ConvertFrom-IsoValue $null | Should -BeNullOrEmpty
        ConvertFrom-IsoValue ' ' | Should -BeNullOrEmpty
    }
}
