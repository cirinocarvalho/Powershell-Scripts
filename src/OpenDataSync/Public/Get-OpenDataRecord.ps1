function Get-OpenDataRecord {
    <#
    .SYNOPSIS
        Retrieves records from a public open data REST API with retry and bearer token support.

    .DESCRIPTION
        Wraps Invoke-RestMethod with the behaviour an unattended job needs: an explicit timeout,
        bounded retries with exponential backoff for transient failures, and a terminating error
        once retries are exhausted. Nothing is suppressed.

        When -RecordPath is supplied the response is drilled into before records are emitted,
        which handles APIs that nest their payload (for example GeoJSON 'features').

    .PARAMETER Uri
        The API endpoint.

    .PARAMETER QueryParameter
        Hashtable appended to the URI as a query string. Values are URL encoded.

    .PARAMETER TokenSecretName
        Name of a secret resolved through Get-AutomationSecret and sent as a bearer token. The
        token is optional: when it cannot be resolved the request is made unauthenticated.

    .PARAMETER Header
        Additional request headers.

    .PARAMETER RecordPath
        Dot separated property path to the collection inside the response, such as 'features'.

    .PARAMETER MaximumRetryCount
        Number of retries after the initial attempt. Defaults to 3.

    .PARAMETER RetryDelaySeconds
        Base delay for exponential backoff. Defaults to 2 seconds.

    .PARAMETER TimeoutSeconds
        Per request timeout. Defaults to 60 seconds.

    .EXAMPLE
        Get-OpenDataRecord -Uri 'https://earthquake.usgs.gov/fdsnws/event/1/query' -QueryParameter @{ format = 'geojson'; limit = 25 } -RecordPath 'features'

        Returns the 25 most recent earthquake features from the USGS public API.

    .OUTPUTS
        System.Management.Automation.PSObject
    #>
    [CmdletBinding()]
    [OutputType([psobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [hashtable]$QueryParameter,

        [string]$TokenSecretName,

        [hashtable]$Header,

        [string]$RecordPath,

        [ValidateRange(0, 10)]
        [int]$MaximumRetryCount = 3,

        [ValidateRange(0, 60)]
        [int]$RetryDelaySeconds = 2,

        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 60
    )

    $requestUri = $Uri
    if ($QueryParameter -and $QueryParameter.Count -gt 0) {
        $pairs = foreach ($key in $QueryParameter.Keys | Sort-Object) {
            '{0}={1}' -f [uri]::EscapeDataString([string]$key), [uri]::EscapeDataString([string]$QueryParameter[$key])
        }
        $separator = if ($Uri.Contains('?')) { '&' } else { '?' }
        $requestUri = $Uri + $separator + ($pairs -join '&')
    }

    $requestHeader = @{ 'User-Agent' = 'OpenDataSync/1.0 (+https://github.com/cirinocarvalho/Powershell-Scripts)' }
    if ($Header) {
        foreach ($key in $Header.Keys) {
            $requestHeader[$key] = $Header[$key]
        }
    }

    if ($TokenSecretName) {
        $token = Get-AutomationSecret -Name $TokenSecretName -AllowMissing
        if ($token) {
            $requestHeader['Authorization'] = "Bearer $token"
        }
        else {
            Write-AutomationLog -Message "No token found for secret '$TokenSecretName'; continuing unauthenticated." -Level Warning
        }
    }

    $attempt = 0
    $response = $null

    while ($true) {
        $attempt++
        try {
            Write-AutomationLog -Message "GET $requestUri (attempt $attempt of $($MaximumRetryCount + 1))" -Level Verbose
            $response = Invoke-RestMethod -Uri $requestUri -Headers $requestHeader -Method Get `
                -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            break
        }
        catch {
            if ($attempt -gt $MaximumRetryCount) {
                throw [System.Net.Http.HttpRequestException]::new(
                    "Request to '$requestUri' failed after $attempt attempt(s): $($_.Exception.Message)",
                    $_.Exception)
            }

            $delay = $RetryDelaySeconds * [Math]::Pow(2, $attempt - 1)
            Write-AutomationLog -Message "Request failed ($($_.Exception.Message)). Retrying in $delay second(s)." -Level Warning
            if ($delay -gt 0) {
                Start-Sleep -Seconds $delay
            }
        }
    }

    $records = $response
    if ($RecordPath) {
        foreach ($segment in $RecordPath.Split('.')) {
            if ($null -eq $records) {
                break
            }

            $property = $records.PSObject.Properties[$segment]
            if (-not $property) {
                throw [System.InvalidOperationException]::new(
                    "Record path segment '$segment' was not found in the response from '$requestUri'.")
            }

            $records = $property.Value
        }
    }

    if ($null -eq $records) {
        Write-AutomationLog -Message "Response from '$requestUri' contained no records." -Level Warning
        return
    }

    Write-Output $records
}
