BeforeAll {
    Import-Module -Name (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1') -Force -ErrorAction Stop
}

Describe 'Assert-SqlIdentifier' {
    It 'brackets a one part identifier' {
        InModuleScope OpenDataSync { Assert-SqlIdentifier -Name 'Assets' } | Should -Be '[Assets]'
    }

    It 'brackets each part of a schema qualified identifier' {
        InModuleScope OpenDataSync { Assert-SqlIdentifier -Name 'dbo.Assets' } | Should -Be '[dbo].[Assets]'
    }

    It 'rejects the injection payload <_>' -ForEach @(
        'Assets];DROP TABLE Users--',
        'Assets Users',
        'a.b.c',
        '1Assets',
        'Assets-Name',
        ''
    ) {
        $payload = $_
        { InModuleScope OpenDataSync -Parameters @{ p = $payload } { Assert-SqlIdentifier -Name $p } } |
            Should -Throw
    }
}

Describe 'New-SqlInsertStatement' {
    It 'builds a parameterized statement' {
        InModuleScope OpenDataSync {
            New-SqlInsertStatement -Table 'dbo.Earthquake' -Column @('Id', 'Magnitude')
        } | Should -Be 'INSERT INTO [dbo].[Earthquake] ([Id], [Magnitude]) VALUES (@Id, @Magnitude);'
    }

    It 'rejects duplicate columns regardless of case' {
        { InModuleScope OpenDataSync { New-SqlInsertStatement -Table 'T' -Column @('Id', 'id') } } |
            Should -Throw -ExpectedMessage '*Duplicate column*'
    }

    It 'rejects an empty column list' {
        { InModuleScope OpenDataSync { New-SqlInsertStatement -Table 'T' -Column @() } } | Should -Throw
    }

    It 'rejects an injected table name' {
        { InModuleScope OpenDataSync { New-SqlInsertStatement -Table 'T];DELETE FROM T--' -Column @('Id') } } |
            Should -Throw
    }
}

Describe 'ConvertTo-SqlParameterSet' {
    It 'maps property names and script blocks' {
        $result = InModuleScope OpenDataSync {
            $record = [pscustomobject]@{ id = 'nc123'; properties = [pscustomobject]@{ mag = 4.2 } }
            ConvertTo-SqlParameterSet -InputObject $record -ColumnMap @{
                Id        = 'id'
                Magnitude = { $_.properties.mag }
            }
        }

        $result['Id'] | Should -Be 'nc123'
        $result['Magnitude'] | Should -Be 4.2
    }

    It 'orders columns deterministically' {
        $keys = InModuleScope OpenDataSync {
            $set = ConvertTo-SqlParameterSet -InputObject ([pscustomobject]@{ a = 1 }) -ColumnMap @{
                Zebra = 'a'; Apple = 'a'; Mango = 'a'
            }
            @($set.Keys)
        }

        $keys | Should -Be @('Apple', 'Mango', 'Zebra')
    }

    It 'converts a missing property to DBNull' {
        $value = InModuleScope OpenDataSync {
            $set = ConvertTo-SqlParameterSet -InputObject ([pscustomobject]@{ a = 1 }) -ColumnMap @{ Missing = 'nope' }
            $set['Missing']
        }

        $value | Should -Be ([System.DBNull]::Value)
    }

    It 'converts a null value to DBNull' {
        $value = InModuleScope OpenDataSync {
            $set = ConvertTo-SqlParameterSet -InputObject ([pscustomobject]@{ a = $null }) -ColumnMap @{ A = 'a' }
            $set['A']
        }

        $value | Should -Be ([System.DBNull]::Value)
    }

    It 'rejects an empty column map' {
        { InModuleScope OpenDataSync {
                ConvertTo-SqlParameterSet -InputObject ([pscustomobject]@{ a = 1 }) -ColumnMap @{}
            } } | Should -Throw '*ColumnMap cannot be empty*'
    }

    It 'surfaces an exception thrown inside a column map script block' {
        { InModuleScope OpenDataSync {
                ConvertTo-SqlParameterSet -InputObject ([pscustomobject]@{ a = 1 }) -ColumnMap @{
                    Bad = { throw 'mapping failure' }
                }
            } } | Should -Throw '*mapping failure*'
    }
}
