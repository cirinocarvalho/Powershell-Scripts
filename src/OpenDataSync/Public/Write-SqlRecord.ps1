function Write-SqlRecord {
    <#
    .SYNOPSIS
        Loads pipeline records into a SQL Server table using parameterized inserts.

    .DESCRIPTION
        Records are projected through a column map, batched, and inserted inside a transaction
        per batch. Identifiers are validated and values are always bound as parameters, so record
        content cannot alter the statement.

        Supports -WhatIf, which reports the statement and row counts without opening a connection.

    .PARAMETER InputObject
        Records to load. Accepts pipeline input.

    .PARAMETER Table
        Destination table, optionally schema qualified, for example 'dbo.Earthquake'.

    .PARAMETER ColumnMap
        Hashtable mapping destination column names to either a source property name or a script
        block that receives the record as $_.

    .PARAMETER ConnectionString
        SQL Server connection string. Defaults to the 'SqlConnectionString' secret.

    .PARAMETER BatchSize
        Rows per transaction. Defaults to 500.

    .EXAMPLE
        $map = @{ Id = 'id'; Magnitude = { $_.properties.mag } }
        $records | Write-SqlRecord -Table dbo.Earthquake -ColumnMap $map -WhatIf

        Shows what would be inserted without connecting to the database.

    .OUTPUTS
        PSCustomObject with RowsRead, RowsWritten and BatchCount.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [psobject]$InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Table,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$ColumnMap,

        [string]$ConnectionString,

        [ValidateRange(1, 10000)]
        [int]$BatchSize = 500
    )

    begin {
        $columns = @($ColumnMap.Keys | Sort-Object)
        $statement = New-SqlInsertStatement -Table $Table -Column $columns

        $resolvedConnectionString = $ConnectionString
        if ([string]::IsNullOrWhiteSpace($resolvedConnectionString) -and -not $WhatIfPreference) {
            $resolvedConnectionString = Get-AutomationSecret -Name 'SqlConnectionString'
        }

        $buffer = [System.Collections.Generic.List[System.Collections.Specialized.OrderedDictionary]]::new()
        $rowsRead = 0
        $rowsWritten = 0
        $batchCount = 0

        $flush = {
            if ($buffer.Count -eq 0) {
                return
            }

            $batchCount++
            $target = "$($buffer.Count) row(s) into $Table"
            if ($PSCmdlet.ShouldProcess($target, 'INSERT')) {
                $written = Invoke-SqlInsertBatch -ConnectionString $resolvedConnectionString `
                    -Statement $statement -ParameterSet $buffer.ToArray()
                $rowsWritten += $written
                Write-AutomationLog -Message "Committed batch $batchCount ($written row(s))." -Level Verbose
            }

            $buffer.Clear()
        }
    }

    process {
        if ($null -eq $InputObject) {
            return
        }

        $rowsRead++
        $buffer.Add((ConvertTo-SqlParameterSet -InputObject $InputObject -ColumnMap $ColumnMap))

        if ($buffer.Count -ge $BatchSize) {
            # Dot sourced so the counters it updates stay in the function scope.
            . $flush
        }
    }

    end {
        . $flush

        Write-AutomationLog -Message "Read $rowsRead record(s); wrote $rowsWritten row(s) to $Table in $batchCount batch(es)."

        [pscustomobject]@{
            Table       = $Table
            Statement   = $statement
            RowsRead    = $rowsRead
            RowsWritten = $rowsWritten
            BatchCount  = $batchCount
        }
    }
}
