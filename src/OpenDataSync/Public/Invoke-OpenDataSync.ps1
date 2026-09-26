function Invoke-OpenDataSync {
    <#
    .SYNOPSIS
        Runs the end to end open data to SQL Server pipeline.

    .DESCRIPTION
        Composes Get-OpenDataRecord and Write-SqlRecord into a single unattended job with
        structured logging and a summary object suitable for a scheduler or CI step. Any failure
        is terminating and non zero, so a scheduled task does not silently report success.

    .PARAMETER Uri
        The open data API endpoint.

    .PARAMETER Table
        Destination SQL Server table.

    .PARAMETER ColumnMap
        Hashtable mapping destination columns to source property names or script blocks.

    .PARAMETER QueryParameter
        Query string values for the API request.

    .PARAMETER RecordPath
        Dot separated path to the record collection inside the response.

    .PARAMETER TokenSecretName
        Optional secret name for a bearer token.

    .PARAMETER ConnectionString
        SQL Server connection string. Defaults to the 'SqlConnectionString' secret.

    .PARAMETER LogPath
        Optional log file path.

    .EXAMPLE
        Invoke-OpenDataSync -Uri 'https://earthquake.usgs.gov/fdsnws/event/1/query' -QueryParameter @{ format = 'geojson'; limit = 10 } -RecordPath features -Table dbo.Earthquake -ColumnMap @{ Id = 'id' } -WhatIf

    .OUTPUTS
        PSCustomObject summarising the run.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '',
        Justification = 'ShouldProcess is delegated to Write-SqlRecord, which performs the state change. Declaring it here lets -WhatIf flow down to that call.')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Table,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$ColumnMap,

        [hashtable]$QueryParameter,

        [string]$RecordPath,

        [string]$TokenSecretName,

        [string]$ConnectionString,

        [string]$LogPath
    )

    $logSplat = @{}
    if ($LogPath) {
        $logSplat['Path'] = $LogPath
    }

    $started = [DateTimeOffset]::UtcNow
    Write-AutomationLog -Message "Starting sync from '$Uri' into '$Table'." @logSplat

    try {
        $fetchSplat = @{ Uri = $Uri }
        foreach ($key in 'QueryParameter', 'RecordPath', 'TokenSecretName') {
            if ($PSBoundParameters.ContainsKey($key) -and $PSBoundParameters[$key]) {
                $fetchSplat[$key] = $PSBoundParameters[$key]
            }
        }

        $records = @(Get-OpenDataRecord @fetchSplat)
        Write-AutomationLog -Message "Retrieved $($records.Count) record(s)." @logSplat

        $writeSplat = @{
            Table     = $Table
            ColumnMap = $ColumnMap
        }
        if ($ConnectionString) {
            $writeSplat['ConnectionString'] = $ConnectionString
        }

        $result = $records | Write-SqlRecord @writeSplat

        $summary = [pscustomobject]@{
            Uri         = $Uri
            Table       = $Table
            RowsRead    = $result.RowsRead
            RowsWritten = $result.RowsWritten
            BatchCount  = $result.BatchCount
            StartedUtc  = $started
            DurationMs  = [int]([DateTimeOffset]::UtcNow - $started).TotalMilliseconds
            Succeeded   = $true
        }

        Write-AutomationLog -Message "Sync completed: $($summary.RowsWritten) row(s) written in $($summary.DurationMs) ms." @logSplat
        return $summary
    }
    catch {
        Write-AutomationLog -Message "Sync failed: $($_.Exception.Message)" -Level Error @logSplat
        throw
    }
}
