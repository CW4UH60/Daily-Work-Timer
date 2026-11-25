
BeforeAll {
    # Pester discovers each test file before it runs it. Load the test harness during
    # the run phase so its setup and helper functions are available to the tests.
    . (Join-Path $PSScriptRoot 'TestSupport.ps1')
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Initialize-WorkTimerTestEnvironment
    Import-WorkTimerFunctions -Name @(
        'Initialize-Storage', 'New-Id', 'New-List', 'Get-HistorySortKey', 'Sort-History',
        'ConvertTo-XmlText', 'Format-XlsxTextCell', 'Format-XlsxNumberCell',
        'Format-XlsxFormulaCell', 'Write-ZipPackage', 'Export-HistoryWorkbook'
    )
}

Describe 'Excel work-history export' {
    It 'writes a readable OOXML workbook with the expected table, headers, formulas, and history data' {
        Set-WorkTimerTestStorage -Root $TestDrive
        $global:History = [Collections.ArrayList]::new()
        $day = [datetime]'2026-07-03T00:00:00'
        [void]$global:History.Add((New-TestRecord -Date '2026-07-03' -SessionStart $day.AddHours(8) -SessionEnd $day.AddHours(16.5) `
            -ActiveSeconds 27000 -ActiveHours 7.5 -InactiveSeconds 3600 `
            -Periods @((New-TestPeriod $day.AddHours(12) $day.AddHours(13) 'Manual')) `
            -Reason 'Ended manually'))

        Export-HistoryWorkbook

        Test-Path -LiteralPath $global:WorkbookPath -PathType Leaf | Should -BeTrue
        $archive = [IO.Compression.ZipFile]::OpenRead($global:WorkbookPath)
        try {
            $names = @($archive.Entries | ForEach-Object FullName)
            $names | Should -Contain '[Content_Types].xml'
            $names | Should -Contain 'xl/workbook.xml'
            $names | Should -Contain 'xl/worksheets/sheet1.xml'
            $names | Should -Contain 'xl/tables/table1.xml'

            $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
            $reader = [IO.StreamReader]::new($sheetEntry.Open(), [Text.Encoding]::UTF8)
            try { $sheet = $reader.ReadToEnd() } finally { $reader.Dispose() }

            $tableEntry = $archive.GetEntry('xl/tables/table1.xml')
            $tableReader = [IO.StreamReader]::new($tableEntry.Open(), [Text.Encoding]::UTF8)
            try { $table = $tableReader.ReadToEnd() } finally { $tableReader.Dispose() }
        }
        finally {
            $archive.Dispose()
        }

        $sheet | Should -Match 'Hours Worked'
        $sheet | Should -Match 'SUBTOTAL\(109,WorkHistory\[Hours Worked\]\)'
        $sheet | Should -Match 'Ended manually'
        $table | Should -Match 'name="WorkHistory"'
        $table | Should -Match 'calculatedColumnFormula'
        $global:ExcelPending | Should -BeFalse
    }

    It 'creates a valid empty workbook table when history has no records' {
        Set-WorkTimerTestStorage -Root (Join-Path $TestDrive 'empty-history')
        $global:History = [Collections.ArrayList]::new()

        Export-HistoryWorkbook

        $archive = [IO.Compression.ZipFile]::OpenRead($global:WorkbookPath)
        try {
            $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
            $reader = [IO.StreamReader]::new($sheetEntry.Open(), [Text.Encoding]::UTF8)
            try { $sheet = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        finally {
            $archive.Dispose()
        }

        $sheet | Should -Match 'A1:H3'
        $sheet | Should -Match 'SUBTOTAL\(109,WorkHistory\[Inactive Hours\]\)'
    }
}
