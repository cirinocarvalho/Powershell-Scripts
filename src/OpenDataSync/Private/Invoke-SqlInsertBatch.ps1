function Invoke-SqlInsertBatch {
    <#
    .SYNOPSIS
        Executes one parameterized INSERT statement for each parameter set inside a transaction.

    .DESCRIPTION
        The whole batch commits or rolls back together, so a partial load never reaches the
        target table. This is the only function in the module that touches a live database,
        which keeps the rest of the pipeline unit testable.

    .PARAMETER ConnectionString
        SQL Server connection string.

    .PARAMETER Statement
        Parameterized statement produced by New-SqlInsertStatement.

    .PARAMETER ParameterSet
        One ordered dictionary of column/value pairs per row.

    .OUTPUTS
        System.Int32. Number of rows affected.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [string]$ConnectionString,

        [Parameter(Mandatory)]
        [string]$Statement,

        [Parameter(Mandatory)]
        [System.Collections.Specialized.OrderedDictionary[]]$ParameterSet,

        [int]$CommandTimeoutSeconds = 30
    )

    $factory = Resolve-SqlClientFactory

    $connection = $factory.CreateConnection()
    $connection.ConnectionString = $ConnectionString

    $transaction = $null
    $affected = 0

    try {
        $connection.Open()
        $transaction = $connection.BeginTransaction()

        foreach ($set in $ParameterSet) {
            $command = $connection.CreateCommand()
            try {
                $command.Transaction = $transaction
                $command.CommandText = $Statement
                $command.CommandTimeout = $CommandTimeoutSeconds

                foreach ($column in $set.Keys) {
                    $parameter = $command.CreateParameter()
                    $parameter.ParameterName = "@$column"
                    $parameter.Value = $set[$column]
                    $null = $command.Parameters.Add($parameter)
                }

                $affected += $command.ExecuteNonQuery()
            }
            finally {
                $command.Dispose()
            }
        }

        $transaction.Commit()
        return $affected
    }
    catch {
        if ($transaction) {
            try {
                $transaction.Rollback()
                Write-AutomationLog -Message 'Transaction rolled back; no rows were committed.' -Level Warning
            }
            catch {
                Write-AutomationLog -Message "Rollback failed: $($_.Exception.Message)" -Level Warning
            }
        }

        throw
    }
    finally {
        if ($transaction) { $transaction.Dispose() }
        $connection.Dispose()
    }
}
