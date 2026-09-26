function ConvertTo-SqlParameterSet {
    <#
    .SYNOPSIS
        Projects a source record into an ordered column/value map using a column map.

    .DESCRIPTION
        The column map keys are destination column names. Each value is either the name of a
        property on the source record, or a script block that receives the record as $_ and
        returns the value. Missing properties and $null both resolve to [DBNull]::Value so the
        insert stays well formed instead of failing at execution time.

    .PARAMETER InputObject
        The source record.

    .PARAMETER ColumnMap
        Hashtable of column name to property name or script block.

    .OUTPUTS
        System.Collections.Specialized.OrderedDictionary
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [psobject]$InputObject,

        [Parameter(Mandatory)]
        [hashtable]$ColumnMap
    )

    if ($ColumnMap.Count -eq 0) {
        throw [System.ArgumentException]::new('ColumnMap cannot be empty.', 'ColumnMap')
    }

    $result = [ordered]@{}

    foreach ($column in $ColumnMap.Keys | Sort-Object) {
        $selector = $ColumnMap[$column]
        $value = $null

        if ($selector -is [scriptblock]) {
            try {
                $value = $InputObject | ForEach-Object -Process $selector
            }
            catch {
                throw "Column map for '$column' threw an exception: $($_.Exception.Message)"
            }
        }
        elseif ($null -ne $InputObject) {
            $property = $InputObject.PSObject.Properties[[string]$selector]
            if ($property) {
                $value = $property.Value
            }
        }

        if ($value -is [object[]]) {
            $value = $value | Select-Object -First 1
        }

        if ($null -eq $value -or $value -is [System.DBNull]) {
            $result[$column] = [System.DBNull]::Value
        }
        else {
            $result[$column] = $value
        }
    }

    return $result
}
