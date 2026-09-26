@{
    RootModule            = 'OpenDataSync.psm1'
    ModuleVersion         = '1.0.0'
    GUID                  = '3f4c2d6e-8a1b-4c5d-9e07-2b6f1a9d4c83'
    Author                = 'Cirino Carvalho'
    Copyright             = '(c) Cirino Carvalho. All rights reserved.'
    Description           = 'Cross-platform PowerShell automation module that pulls records from a public open data API and loads them into SQL Server with parameterized inserts, retries and structured logging.'

    PowerShellVersion     = '7.2'
    CompatiblePSEditions  = @('Core')

    FunctionsToExport     = @(
        'Get-AutomationSecret',
        'Get-OpenDataRecord',
        'Invoke-OpenDataSync',
        'Write-SqlRecord'
    )
    CmdletsToExport       = @()
    VariablesToExport     = @()
    AliasesToExport       = @()

    PrivateData           = @{
        PSData = @{
            Tags       = @('automation', 'etl', 'opendata', 'sql', 'crossplatform')
            LicenseUri = 'https://github.com/cirinocarvalho/OpenDataSync/blob/master/LICENSE'
            ProjectUri = 'https://github.com/cirinocarvalho/OpenDataSync'
        }
    }
}
