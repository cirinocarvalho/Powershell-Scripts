# Powershell-Scripts

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows)

A collection of PowerShell scripts for automating everyday Windows tasks, from
document conversion to REST API data ingestion.

## Scripts

| Script | Description |
| --- | --- |
| [`src/Windows Task/PDF2TIFF.ps1`](src/Windows%20Task/PDF2TIFF.ps1) | Converts PDF files to a single multi-page TIFF using Ghostscript, then archives the processed PDFs. |
| [`src/API/BIGBELLY_CLEAN_ASSETS.ps1`](src/API/BIGBELLY_CLEAN_ASSETS.ps1) | Pulls asset data from the BigBelly REST API and inserts it into a SQL Server database, with logging. |

## Requirements

- Windows PowerShell 5.1 or later
- [Ghostscript](https://www.ghostscript.com/) — required by `PDF2TIFF.ps1`
- SQL Server access — required by `BIGBELLY_CLEAN_ASSETS.ps1`

## Usage

Clone the repository and run a script from a PowerShell prompt:

```powershell
git clone https://github.com/cirinocarvalho/Powershell-Scripts.git
cd Powershell-Scripts
.\src\API\BIGBELLY_CLEAN_ASSETS.ps1
```

Open each script and review the variables in the **Declarations** section
(paths, connection strings, API credentials) before running it.

If execution is blocked by policy, allow local scripts for the current session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## License

Released under the [MIT License](LICENSE).
