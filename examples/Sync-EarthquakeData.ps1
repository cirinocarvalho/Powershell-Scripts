#Requires -Version 7.2
<#
.SYNOPSIS
    Loads recent earthquakes from the USGS public open data API into SQL Server.

.DESCRIPTION
    A complete, runnable example of the OpenDataSync module. The USGS feed needs no API key,
    so the example works anywhere, while the connection string still comes from a secret.

    Create the target table first with sql/Earthquake.sql, then provide the connection string:

        $env:OPENDATASYNC_SQLCONNECTIONSTRING = 'Server=localhost;Database=OpenData;...'

    or store it once with SecretManagement:

        Set-Secret -Name SqlConnectionString -Secret 'Server=localhost;Database=OpenData;...'

.PARAMETER MinimumMagnitude
    Smallest magnitude to import.

.PARAMETER Days
    How many days back to query.

.PARAMETER Table
    Destination table.

.PARAMETER LogPath
    Optional log file.

.EXAMPLE
    ./examples/Sync-EarthquakeData.ps1 -WhatIf

    Shows what would be inserted without connecting to a database.

.EXAMPLE
    ./examples/Sync-EarthquakeData.ps1 -MinimumMagnitude 5 -Days 7 -LogPath ./logs/earthquake.log
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(0, 10)]
    [double]$MinimumMagnitude = 4.5,

    [ValidateRange(1, 30)]
    [int]$Days = 1,

    [string]$Table = 'dbo.Earthquake',

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath '../src/OpenDataSync/OpenDataSync.psd1') -Force

# Keys are destination columns. Values are either a source property name or a script block
# that receives the record as $_, which keeps nested GeoJSON access readable.
$columnMap = @{
    EventId      = 'id'
    Magnitude    = { $_.properties.mag }
    Place        = { $_.properties.place }
    EventTimeUtc = { [DateTimeOffset]::FromUnixTimeMilliseconds([long]$_.properties.time).UtcDateTime }
    Longitude    = { $_.geometry.coordinates[0] }
    Latitude     = { $_.geometry.coordinates[1] }
    DepthKm      = { $_.geometry.coordinates[2] }
    DetailUrl    = { $_.properties.url }
}

$syncParameters = @{
    Uri            = 'https://earthquake.usgs.gov/fdsnws/event/1/query'
    QueryParameter = @{
        format    = 'geojson'
        starttime = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-dd')
        minmagnitude = $MinimumMagnitude
        orderby   = 'time'
    }
    RecordPath     = 'features'
    Table          = $Table
    ColumnMap      = $columnMap
}

if ($LogPath) {
    $syncParameters['LogPath'] = $LogPath
}

$summary = Invoke-OpenDataSync @syncParameters
$summary | Format-List
