BeforeAll {
    Import-Module -Name (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1') -Force -ErrorAction Stop

    function Get-TestRecord {
        param([int]$Count)
        1..$Count | ForEach-Object {
            [pscustomobject]@{
                id         = "nc$_"
                properties = [pscustomobject]@{ mag = $_ / 2 ; place = "Place $_" }
            }
        }
    }

    $script:Map = @{
        Id        = 'id'
        Magnitude = { $_.properties.mag }
        Place     = { $_.properties.place }
    }
}

Describe 'Write-SqlRecord' {
    BeforeEach {
        Mock -ModuleName OpenDataSync Invoke-SqlInsertBatch {
            $script:Batches += , $ParameterSet
            $ParameterSet.Count
        }
        $script:Batches = @()
    }

    It 'writes every record and reports the totals' {
        $result = Get-TestRecord -Count 5 |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -InformationAction SilentlyContinue

        $result.RowsRead | Should -Be 5
        $result.RowsWritten | Should -Be 5
        $result.BatchCount | Should -Be 1
    }

    It 'splits work into batches of -BatchSize' {
        $result = Get-TestRecord -Count 7 |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -BatchSize 3 -InformationAction SilentlyContinue

        $result.BatchCount | Should -Be 3
        $result.RowsWritten | Should -Be 7
        $script:Batches.Count | Should -Be 3
        $script:Batches[0].Count | Should -Be 3
        $script:Batches[2].Count | Should -Be 1
    }

    It 'passes a parameterized statement to the database layer' {
        $result = Get-TestRecord -Count 1 |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -InformationAction SilentlyContinue

        $result.Statement | Should -Be 'INSERT INTO [dbo].[Earthquake] ([Id], [Magnitude], [Place]) VALUES (@Id, @Magnitude, @Place);'
        $result.Statement | Should -Not -Match 'nc1'
    }

    It 'binds record content as values rather than statement text' {
        $null = [pscustomobject]@{ id = "x'); DROP TABLE Users--"; properties = [pscustomobject]@{ mag = 1; place = 'p' } } |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -InformationAction SilentlyContinue

        $script:Batches[0][0]['Id'] | Should -Be "x'); DROP TABLE Users--"
    }

    It 'does not touch the database under -WhatIf' {
        $result = Get-TestRecord -Count 4 |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -WhatIf -InformationAction SilentlyContinue

        $result.RowsRead | Should -Be 4
        $result.RowsWritten | Should -Be 0
        Should -Invoke -ModuleName OpenDataSync Invoke-SqlInsertBatch -Times 0 -Exactly
    }

    It 'resolves the connection string from the secret store when not supplied' {
        Mock -ModuleName OpenDataSync Get-AutomationSecret { 'from-secret-store' }

        $null = Get-TestRecord -Count 1 |
            Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -InformationAction SilentlyContinue

        Should -Invoke -ModuleName OpenDataSync Get-AutomationSecret -Times 1 -Exactly `
            -ParameterFilter { $Name -eq 'SqlConnectionString' }
        Should -Invoke -ModuleName OpenDataSync Invoke-SqlInsertBatch -Times 1 -Exactly `
            -ParameterFilter { $ConnectionString -eq 'from-secret-store' }
    }

    It 'handles an empty pipeline without calling the database' {
        $result = @() | Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' -InformationAction SilentlyContinue

        $result.RowsRead | Should -Be 0
        $result.BatchCount | Should -Be 0
        Should -Invoke -ModuleName OpenDataSync Invoke-SqlInsertBatch -Times 0 -Exactly
    }

    It 'fails fast on an invalid table name before reading the pipeline' {
        $records = Get-TestRecord -Count 1
        $act = { $records | Write-SqlRecord -Table 'dbo.Bad];DROP TABLE x--' -ColumnMap $script:Map -ConnectionString 'fake' }

        $act | Should -Throw

        Should -Invoke -ModuleName OpenDataSync Invoke-SqlInsertBatch -Times 0 -Exactly
    }

    It 'propagates a database failure instead of swallowing it' {
        Mock -ModuleName OpenDataSync Invoke-SqlInsertBatch { throw 'deadlock victim' }
        $records = Get-TestRecord -Count 1
        $act = { $records | Write-SqlRecord -Table 'dbo.Earthquake' -ColumnMap $script:Map -ConnectionString 'fake' }

        $act | Should -Throw -ExpectedMessage '*deadlock victim*'
    }
}

Describe 'Invoke-OpenDataSync' {
    BeforeEach {
        Mock -ModuleName OpenDataSync Invoke-SqlInsertBatch { $ParameterSet.Count }
        Mock -ModuleName OpenDataSync Get-OpenDataRecord { Get-TestRecord -Count 3 }
    }

    It 'returns a summary for a successful run' {
        $summary = Invoke-OpenDataSync -Uri 'https://example.test/api' -Table 'dbo.Earthquake' `
            -ColumnMap $script:Map -ConnectionString 'fake' -InformationAction SilentlyContinue

        $summary.Succeeded | Should -BeTrue
        $summary.RowsRead | Should -Be 3
        $summary.RowsWritten | Should -Be 3
        $summary.Uri | Should -Be 'https://example.test/api'
        $summary.DurationMs | Should -BeGreaterOrEqual 0
    }

    It 'forwards fetch options to Get-OpenDataRecord' {
        $null = Invoke-OpenDataSync -Uri 'https://example.test/api' -Table 'dbo.Earthquake' `
            -ColumnMap $script:Map -ConnectionString 'fake' -RecordPath 'features' `
            -QueryParameter @{ limit = 3 } -InformationAction SilentlyContinue

        Should -Invoke -ModuleName OpenDataSync Get-OpenDataRecord -Times 1 -Exactly `
            -ParameterFilter { $RecordPath -eq 'features' -and $QueryParameter.limit -eq 3 }
    }

    It 'rethrows when the fetch fails so a scheduled job reports failure' {
        Mock -ModuleName OpenDataSync Get-OpenDataRecord { throw 'API unreachable' }

        { Invoke-OpenDataSync -Uri 'https://example.test/api' -Table 'dbo.Earthquake' `
                -ColumnMap $script:Map -ConnectionString 'fake' -ErrorAction Stop } |
            Should -Throw -ExpectedMessage '*API unreachable*'
    }

    It 'propagates -WhatIf down to the write stage so nothing is inserted' {
        $summary = Invoke-OpenDataSync -Uri 'https://example.test/api' -Table 'dbo.Earthquake' `
            -ColumnMap $script:Map -ConnectionString 'fake' -WhatIf -InformationAction SilentlyContinue

        $summary.RowsRead | Should -Be 3
        $summary.RowsWritten | Should -Be 0
        Should -Invoke -ModuleName OpenDataSync Invoke-SqlInsertBatch -Times 0 -Exactly
    }

    It 'writes to a log file when -LogPath is supplied' {        $logPath = Join-Path -Path $TestDrive -ChildPath 'logs/sync.log'

        $null = Invoke-OpenDataSync -Uri 'https://example.test/api' -Table 'dbo.Earthquake' `
            -ColumnMap $script:Map -ConnectionString 'fake' -LogPath $logPath -InformationAction SilentlyContinue

        Test-Path -LiteralPath $logPath | Should -BeTrue
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'Starting sync'
    }
}
