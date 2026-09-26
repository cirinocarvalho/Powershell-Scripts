function Assert-SqlIdentifier {
    <#
    .SYNOPSIS
        Validates a SQL Server identifier and returns it wrapped in brackets.

    .DESCRIPTION
        Table and column names cannot be passed as command parameters, so they are concatenated
        into the statement text. This function is the guard that keeps that concatenation safe:
        only plain identifiers are accepted, and a one or two part name is bracket quoted.

    .PARAMETER Name
        A one part ('Assets') or two part ('dbo.Assets') identifier.

    .OUTPUTS
        System.String. The bracket quoted identifier.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw [System.ArgumentException]::new('A SQL identifier cannot be null or empty.', 'Name')
    }

    $parts = $Name.Split('.')
    if ($parts.Count -gt 2) {
        throw [System.ArgumentException]::new(
            "SQL identifier '$Name' is not valid. Use 'Object' or 'Schema.Object'.", 'Name')
    }

    $pattern = '^[A-Za-z_][A-Za-z0-9_]*$'
    foreach ($part in $parts) {
        if ($part -notmatch $pattern) {
            throw [System.ArgumentException]::new(
                "SQL identifier part '$part' is not valid. Letters, digits and underscores only, and it cannot start with a digit.",
                'Name')
        }
    }

    return ($parts | ForEach-Object { "[$_]" }) -join '.'
}
