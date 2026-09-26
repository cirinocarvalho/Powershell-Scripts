function New-SqlInsertStatement {
    <#
    .SYNOPSIS
        Builds a parameterized INSERT statement for the given table and columns.

    .DESCRIPTION
        Pure string building with no database dependency, which keeps it fully unit testable.
        Values are always represented as named parameters so record content can never be
        interpreted as SQL.

    .PARAMETER Table
        Target table, optionally schema qualified.

    .PARAMETER Column
        Ordered list of column names.

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Builds a string and changes no state; the New verb describes the returned object.')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$Table,

        [Parameter(Mandatory)]
        [string[]]$Column
    )

    if ($Column.Count -eq 0) {
        throw [System.ArgumentException]::new('At least one column is required.', 'Column')
    }

    $duplicates = $Column | Group-Object -Property { $_.ToLowerInvariant() } | Where-Object Count -GT 1
    if ($duplicates) {
        throw [System.ArgumentException]::new(
            "Duplicate column name(s): $(($duplicates.Name | Sort-Object) -join ', ').", 'Column')
    }

    $safeTable = Assert-SqlIdentifier -Name $Table
    $safeColumns = foreach ($name in $Column) { Assert-SqlIdentifier -Name $name }
    $parameters = foreach ($name in $Column) { "@$name" }

    return 'INSERT INTO {0} ({1}) VALUES ({2});' -f $safeTable, ($safeColumns -join ', '), ($parameters -join ', ')
}
