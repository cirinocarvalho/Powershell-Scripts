#Requires -Version 7.2

Set-StrictMode -Version Latest

$script:ModuleRoot = $PSScriptRoot

foreach ($folder in 'Private', 'Public') {
    $path = Join-Path -Path $PSScriptRoot -ChildPath $folder
    if (-not (Test-Path -LiteralPath $path)) {
        continue
    }

    foreach ($file in Get-ChildItem -LiteralPath $path -Filter '*.ps1' -File | Sort-Object -Property Name) {
        try {
            . $file.FullName
        }
        catch {
            throw "Failed to import '$($file.FullName)': $($_.Exception.Message)"
        }
    }
}

$public = Get-ChildItem -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath 'Public') -Filter '*.ps1' -File |
    Select-Object -ExpandProperty BaseName

Export-ModuleMember -Function $public
