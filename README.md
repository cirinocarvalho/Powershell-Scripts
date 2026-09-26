# OpenDataSync

[![CI](https://github.com/cirinocarvalho/OpenDataSync/actions/workflows/ci.yml/badge.svg)](https://github.com/cirinocarvalho/OpenDataSync/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-7.2%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-success)](#requirements)

A small, production shaped **workflow automation** showcase: `OpenDataSync` is a PowerShell 7
module that pulls records from a public open data API and loads them into SQL Server, with the
things an unattended job actually needs — secrets outside the script, retries, transactions,
real error handling, and tests that run on every push.

## Why this exists

This repository began as a set of typical one-off Windows task scripts: hardcoded paths,
`$ErrorActionPreference = "SilentlyContinue"`, credentials in source, and no tests. They worked,
but they were not something you could hand to a team. This project is the same kind of work
rebuilt the way it should be shipped.

| | Typical one-off script | `OpenDataSync` |
| --- | --- | --- |
| Platform | Windows PowerShell 5.1 | PowerShell 7.2+ on Windows, macOS, Linux |
| Secrets | Embedded in the script | Environment variables or `SecretManagement` |
| Errors | `SilentlyContinue`, failures pass silently | Terminating errors, bounded retries, non-zero exit |
| Writes | String concatenated SQL | Parameterized inserts inside a transaction |
| Safety | None | `-WhatIf` dry run, identifier validation |
| Tests | None | 64 Pester tests + PSScriptAnalyzer in CI on 3 operating systems |

## Quick start

```powershell
git clone https://github.com/cirinocarvalho/OpenDataSync.git
cd OpenDataSync
Import-Module ./src/OpenDataSync/OpenDataSync.psd1

# Dry run against the live USGS open data feed. No database and no secrets needed.
./examples/Sync-EarthquakeData.ps1 -WhatIf
```

To load data for real, create the table and supply a connection string:

```powershell
# 1. Create the destination table (sql/Earthquake.sql)
# 2. Provide the connection string as a secret, never in the script:
$env:OPENDATASYNC_SQLCONNECTIONSTRING = 'Server=localhost;Database=OpenData;Integrated Security=True;'

./examples/Sync-EarthquakeData.ps1 -MinimumMagnitude 5 -Days 7 -LogPath ./logs/earthquake.log
```

## Commands

| Command | Purpose |
| --- | --- |
| `Get-AutomationSecret` | Resolves a secret from an environment variable, then a `SecretManagement` vault. |
| `Get-OpenDataRecord` | Calls a REST API with a timeout, exponential backoff retries, and optional bearer token. |
| `Write-SqlRecord` | Loads pipeline records into a table using parameterized, batched, transactional inserts. |
| `Invoke-OpenDataSync` | Runs the whole fetch-to-load pipeline and returns a run summary. |

Every command ships comment based help: `Get-Help Invoke-OpenDataSync -Full`.

### Secrets

Nothing sensitive belongs in source control. `Get-AutomationSecret` checks the environment
first, which is what CI runners and containers provide, then falls back to a vault:

```powershell
# CI or container
$env:OPENDATASYNC_SQLCONNECTIONSTRING = '...'

# Workstation
Set-Secret -Name SqlConnectionString -Secret '...'
```

A secret named `SqlConnectionString` maps to `OPENDATASYNC_SQLCONNECTIONSTRING`. Missing secrets
raise an actionable error that names the exact variable to set, unless you pass `-AllowMissing`.

### Mapping records to columns

`ColumnMap` keys are destination columns. Values are either a source property name or a script
block receiving the record as `$_`, which keeps nested payloads readable:

```powershell
$map = @{
    EventId      = 'id'
    Magnitude    = { $_.properties.mag }
    EventTimeUtc = { [DateTimeOffset]::FromUnixTimeMilliseconds([long]$_.properties.time).UtcDateTime }
}

Get-OpenDataRecord -Uri $uri -RecordPath features |
    Write-SqlRecord -Table dbo.Earthquake -ColumnMap $map
```

Missing properties and `$null` become `DBNull`, so a sparse payload does not break the load.

### Safety

Table and column names are validated against a strict identifier pattern and bracket quoted;
anything else is rejected before a connection is opened. Record values are always bound as
command parameters, so payload content can never be executed as SQL. Each batch runs in a
transaction and rolls back as a unit, and `-WhatIf` reports the statement and row counts
without touching the database.

## Requirements

- [PowerShell 7.2+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)
- `SqlServer` module for live database writes: `Install-Module SqlServer -Scope CurrentUser`
  (not needed for `-WhatIf` runs or the test suite)
- Optional: `Microsoft.PowerShell.SecretManagement` for vault backed secrets

## Development

```powershell
Install-Module Pester, PSScriptAnalyzer -Scope CurrentUser

./build.ps1 -Task All        # lint + test, exactly what CI runs
./build.ps1 -Task Test
./build.ps1 -Task Analyze
```

CI runs the same `build.ps1` on `ubuntu-latest`, `windows-latest` and `macos-latest`, and fails
on any analyzer finding, any failing test, or code coverage below 80%. The suite needs no
database or network: HTTP calls are mocked, the ADO.NET layer is driven through duck typed
fakes that verify commit, rollback and disposal, and the SQL statement builder and column
mapper are pure functions tested directly.

## Repository layout

```
src/OpenDataSync/     # the module (Public/ exported, Private/ internal)
tests/                # Pester suite
examples/             # runnable end-to-end example
sql/                  # destination table DDL
build.ps1             # lint + test entry point used locally and in CI
```

## License

Released under the [MIT License](LICENSE).
