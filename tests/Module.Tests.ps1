BeforeAll {
    $script:ModulePath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1'
    Import-Module -Name $script:ModulePath -Force -ErrorAction Stop
}

Describe 'Module contract' {
    It 'has a valid manifest' {
        { Test-ModuleManifest -Path $script:ModulePath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'requires PowerShell 7.2 or later and targets Core' {
        $manifest = Import-PowerShellDataFile -Path $script:ModulePath
        $manifest.PowerShellVersion | Should -Be '7.2'
        $manifest.CompatiblePSEditions | Should -Contain 'Core'
    }

    It 'exports exactly the documented public functions' {
        $expected = @('Get-AutomationSecret', 'Get-OpenDataRecord', 'Invoke-OpenDataSync', 'Write-SqlRecord')
        $actual = (Get-Command -Module OpenDataSync -CommandType Function).Name | Sort-Object
        $actual | Should -Be $expected
    }

    It 'does not leak private helpers' {
        Get-Command -Name 'Invoke-SqlInsertBatch' -ErrorAction Ignore | Should -BeNullOrEmpty
    }

    It 'gives <_> comment based help with an example' -ForEach @(
        'Get-AutomationSecret', 'Get-OpenDataRecord', 'Invoke-OpenDataSync', 'Write-SqlRecord'
    ) {
        $help = Get-Help -Name $_ -ErrorAction Stop
        $help.Synopsis | Should -Not -BeNullOrEmpty
        @($help.Examples.Example).Count | Should -BeGreaterThan 0
    }

    It 'contains no plain text secrets in the module source' {
        $files = Get-ChildItem -Path (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src') -Recurse -Filter '*.ps1'
        foreach ($file in $files) {
            $content = Get-Content -LiteralPath $file.FullName -Raw
            $content | Should -Not -Match 'Password\s*=\s*["''][^"'']+["'']'
            $content | Should -Not -Match 'ConvertTo-SecureString\s+.*-AsPlainText\s+.*["''][^"'']{8,}["'']'
        }
    }
}
