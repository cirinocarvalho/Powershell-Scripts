BeforeAll {
    Import-Module -Name (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1') -Force -ErrorAction Stop
}

Describe 'Get-AutomationSecret' {
    AfterEach {
        Remove-Item -Path 'Env:OPENDATASYNC_SQLCONNECTIONSTRING' -ErrorAction Ignore
        Remove-Item -Path 'Env:OPENDATASYNC_SQL_CONNECTION_STRING' -ErrorAction Ignore
        Remove-Item -Path 'Env:CUSTOM_TOKEN_VAR' -ErrorAction Ignore
    }

    It 'resolves a secret from the derived environment variable name' {
        $env:OPENDATASYNC_SQLCONNECTIONSTRING = 'Server=tcp:db;Database=open;'
        Get-AutomationSecret -Name 'SqlConnectionString' | Should -Be 'Server=tcp:db;Database=open;'
    }

    It 'uppercases and normalises punctuation when deriving the variable name' {
        $env:OPENDATASYNC_SQL_CONNECTION_STRING = 'derived'
        Get-AutomationSecret -Name 'sql-connection.string' | Should -Be 'derived'
    }

    It 'prefers an explicit environment variable name' {
        $env:OPENDATASYNC_SQLCONNECTIONSTRING = 'derived'
        $env:CUSTOM_TOKEN_VAR = 'explicit'
        Get-AutomationSecret -Name 'SqlConnectionString' -EnvironmentVariable 'CUSTOM_TOKEN_VAR' |
            Should -Be 'explicit'
    }

    It 'ignores an environment variable that is only whitespace' {
        $env:OPENDATASYNC_SQLCONNECTIONSTRING = '   '
        { Get-AutomationSecret -Name 'SqlConnectionString' } | Should -Throw
    }

    It 'throws an actionable error when the secret is missing' {
        { Get-AutomationSecret -Name 'SqlConnectionString' } |
            Should -Throw -ExpectedMessage '*OPENDATASYNC_SQLCONNECTIONSTRING*'
    }

    It 'returns null instead of throwing when -AllowMissing is used' {
        Get-AutomationSecret -Name 'ApiToken' -AllowMissing | Should -BeNullOrEmpty
    }

    It 'rejects an empty secret name' {
        { Get-AutomationSecret -Name '' } | Should -Throw
    }
}
