function Get-AutomationSecret {
    <#
    .SYNOPSIS
        Resolves a secret from an environment variable or Microsoft.PowerShell.SecretManagement.

    .DESCRIPTION
        Secrets are never stored in the repository. This function looks in the environment first,
        which is what CI systems and containers provide, and falls back to a SecretManagement
        vault for interactive or workstation use.

        The environment variable name defaults to the secret name uppercased with non
        alphanumeric characters replaced by underscores and an OPENDATASYNC_ prefix, so a secret
        named 'SqlConnectionString' maps to OPENDATASYNC_SQLCONNECTIONSTRING.

    .PARAMETER Name
        Logical secret name, also used as the SecretManagement secret name.

    .PARAMETER EnvironmentVariable
        Explicit environment variable to read instead of the derived name.

    .PARAMETER Vault
        Optional SecretManagement vault name. When omitted the registered default vault is used.

    .PARAMETER AllowMissing
        Return $null instead of throwing when the secret cannot be found.

    .EXAMPLE
        $cs = Get-AutomationSecret -Name SqlConnectionString

        Reads OPENDATASYNC_SQLCONNECTIONSTRING, then falls back to the default secret vault.

    .EXAMPLE
        $token = Get-AutomationSecret -Name ApiToken -AllowMissing

        Returns $null when no token is configured, which lets a caller treat it as optional.

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [string]$EnvironmentVariable,

        [string]$Vault,

        [switch]$AllowMissing
    )

    $variableName = if ($PSBoundParameters.ContainsKey('EnvironmentVariable') -and $EnvironmentVariable) {
        $EnvironmentVariable
    }
    else {
        'OPENDATASYNC_' + ($Name -replace '[^A-Za-z0-9]', '_').ToUpperInvariant()
    }

    $value = [Environment]::GetEnvironmentVariable($variableName)
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        Write-Verbose -Message "Secret '$Name' resolved from environment variable '$variableName'."
        return $value
    }

    if (Get-Command -Name Get-Secret -Module Microsoft.PowerShell.SecretManagement -ErrorAction Ignore) {
        try {
            $parameters = @{
                Name        = $Name
                AsPlainText = $true
                ErrorAction = 'Stop'
            }
            if ($Vault) {
                $parameters['Vault'] = $Vault
            }

            $secret = Get-Secret @parameters
            if (-not [string]::IsNullOrWhiteSpace($secret)) {
                Write-Verbose -Message "Secret '$Name' resolved from SecretManagement."
                return $secret
            }
        }
        catch {
            Write-Verbose -Message "SecretManagement lookup for '$Name' failed: $($_.Exception.Message)"
        }
    }

    if ($AllowMissing) {
        Write-Verbose -Message "Secret '$Name' was not found; continuing because -AllowMissing was specified."
        return $null
    }

    throw [System.InvalidOperationException]::new(
        "Secret '$Name' was not found. Set the environment variable '$variableName' or store it with: Set-Secret -Name '$Name'.")
}
