function Resolve-SqlClientFactory {
    <#
    .SYNOPSIS
        Locates a SqlClient provider factory that works on Windows, macOS and Linux.

    .DESCRIPTION
        PowerShell 7 does not ship a SQL Server client. This resolves Microsoft.Data.SqlClient,
        importing the SqlServer module if the types are not already loaded, and fails with an
        actionable message instead of a NullReferenceException deep inside the pipeline.

    .OUTPUTS
        System.Data.Common.DbProviderFactory
    #>
    [CmdletBinding()]
    [OutputType([System.Data.Common.DbProviderFactory])]
    param()

    if ($script:SqlClientFactory) {
        return $script:SqlClientFactory
    }

    $typeName = 'Microsoft.Data.SqlClient.SqlClientFactory'

    if (-not ($typeName -as [type])) {
        Write-Verbose -Message 'Microsoft.Data.SqlClient is not loaded. Importing the SqlServer module.'
        try {
            Import-Module -Name SqlServer -ErrorAction Stop -Verbose:$false
        }
        catch {
            throw [System.InvalidOperationException]::new(
                "Microsoft.Data.SqlClient is not available. Install it with: Install-Module SqlServer -Scope CurrentUser. Underlying error: $($_.Exception.Message)",
                $_.Exception)
        }
    }

    $factoryType = $typeName -as [type]
    if (-not $factoryType) {
        throw [System.InvalidOperationException]::new(
            "The SqlServer module was imported but '$typeName' could not be found.")
    }

    $script:SqlClientFactory = $factoryType::Instance
    return $script:SqlClientFactory
}
