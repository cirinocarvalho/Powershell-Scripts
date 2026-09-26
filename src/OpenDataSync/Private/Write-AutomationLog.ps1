function Write-AutomationLog {
    <#
    .SYNOPSIS
        Writes a structured, timestamped log line to the PowerShell streams and an optional file.

    .DESCRIPTION
        Replaces the ad hoc Log-Write helpers used by the legacy scripts. Output goes to the
        stream that matches the severity so that callers keep full control through the standard
        preference variables, and nothing is swallowed silently.

    .PARAMETER Message
        The text to log.

    .PARAMETER Level
        Severity of the entry. Error entries are written to the error stream as non-terminating
        errors so the caller decides whether to continue.

    .PARAMETER Path
        Optional log file. The parent directory is created when it does not exist. A failure to
        write the file is surfaced as a warning rather than aborting the automation run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('Information', 'Warning', 'Error', 'Verbose')]
        [string]$Level = 'Information',

        [string]$Path
    )

    process {
        $timestamp = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $line = '{0} [{1}] {2}' -f $timestamp, $Level.ToUpperInvariant(), $Message

        switch ($Level) {
            'Warning' { Write-Warning -Message $line }
            'Error' { Write-Error -Message $line -ErrorAction Continue }
            'Verbose' { Write-Verbose -Message $line }
            default { Write-Information -MessageData $line -InformationAction Continue }
        }

        if (-not $PSBoundParameters.ContainsKey('Path') -or [string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        try {
            $directory = Split-Path -Path $Path -Parent
            if ($directory -and -not (Test-Path -LiteralPath $directory)) {
                $null = New-Item -Path $directory -ItemType Directory -Force -ErrorAction Stop
            }

            Add-Content -LiteralPath $Path -Value $line -Encoding utf8 -ErrorAction Stop
        }
        catch {
            Write-Warning -Message "Unable to write to log file '$Path': $($_.Exception.Message)"
        }
    }
}
