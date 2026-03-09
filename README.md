# Google Takeout Metadata Restorer

**[日本語版 README](README_ja.md)**

Restore EXIF metadata (date, GPS, description, title) from Google Takeout `.supplemental-metadata.json` files back into your photos and videos.
This script helps keep chronological order and location info when you migrate media to other photo apps.

## Overview
Google Takeout exports metadata separately from media files. Many photo tools ignore these JSON sidecar files, so imported media can lose capture dates and GPS information.
This project applies JSON metadata back into media files through [ExifTool](https://exiftool.org/).

### Target Users
- Users migrating photos/videos from Google Takeout.
- Users who want to preserve original timestamp and location metadata.

## Features
- Multi-step filename matching for truncated Google Takeout names.
- Parallel ExifTool execution with `-Threads` (`1` to `32`).
- Optional `YYYY/MM/` output hierarchy via `-OutputPath`.
- File-type detection by magic bytes and extension auto-fix.
- Unicode-safe matching for Japanese/CJK filenames.
- Dry-run preview with `-WhatIf`.
- CSV log output and failure report output.
- UI language switch: Japanese or English with `-Language`.

## Requirements
- PowerShell 5.1+ (Windows) or PowerShell 7+.
- [ExifTool](https://exiftool.org/).

## Status
- Last README update: 2026-03-09
- Supported runtime: PowerShell 5.1+ / ExifTool stable release

## Installation
1. Clone this repository.
2. Install ExifTool.

```powershell
git clone https://github.com/cyrne1-7208/google-takeout-metadata-restorer.git
cd google-takeout-metadata-restorer
```

ExifTool installation examples:

```powershell
# Windows
winget install exiftool

# macOS
brew install exiftool

# Ubuntu/Debian
sudo apt install libimage-exiftool-perl
```

## Quick Start
Run a dry-run first to verify matching without modifying files.

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -WhatIf -Language en
```

Expected output:

```text
The script prints matching statistics, planned updates, and a failure summary.
No media files are modified in WhatIf mode.
```

## Usage
In-place update:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos"
```

Write output to year/month folders:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos"
```

Parallel processing:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos" -Threads 8
```

If PowerShell blocks execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

## Configuration
| Parameter | Required | Default | Description |
|---|---|---|---|
| `-PhotosPath` | Yes | None | Google Takeout media root folder |
| `-Extensions` | No | 15 media extensions | Target media extensions |
| `-ExifToolPath` | No | `exiftool` | ExifTool executable path |
| `-WhatIf` | No | `false` | Dry-run mode (no file writes) |
| `-NoBackup` | No | `false` | Skip `_original` backup file creation |
| `-OriginalFileAction` | No | `Keep` | `_original` handling: `Keep` / `Rename` / `Delete` |
| `-OutputPath` | No | Empty | Output root with `YYYY/MM/` hierarchy |
| `-LogFile` | No | `restore-metadata-log.csv` | CSV log output path |
| `-PrefixMatchChars` | No | `20` | Prefix length for filename matching |
| `-TimeToleranceSeconds` | No | `86400` | Timestamp-near matching tolerance (seconds) |
| `-Threads` | No | `1` | Parallel workers (`1` to `32`) |
| `-Language` | No | `ja` | Console/report language: `ja` or `en` |

## Restored Metadata
| Metadata | EXIF tags |
|---|---|
| Capture date/time | `DateTimeOriginal`, `CreateDate` |
| Modified date/time | `ModifyDate`, `FileModifyDate` |
| GPS | `GPSLatitude`, `GPSLongitude`, `GPSAltitude` (+ refs) |
| Description | `ImageDescription`, `XPComment` |
| Title | `Title`, `XPTitle` |

## Project Structure
| Path | Purpose |
|---|---|
| `restore_metadata.ps1` | Main metadata restoration script |
| `README.md` | English documentation |
| `README_ja.md` | Japanese documentation |
| `LICENSE` | MIT license text |

## Testing
There is no automated test suite in this repository.
Validate changes with a small Takeout sample and run `-WhatIf` before actual execution.

## Troubleshooting
- `Execution policy` error: run `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process` in the current shell.
- `exiftool` not found: set `-ExifToolPath` to the full path of `exiftool.exe`.
- Unsure about side effects: run with `-WhatIf` first and review the CSV/failure report.
- Need English output: add `-Language en`.

## Acknowledgements
- Base implementation: [pfilbin90/google-takeout-metadata-restorer](https://github.com/pfilbin90/google-takeout-metadata-restorer)
- Metadata tool: [ExifTool](https://exiftool.org/)

## AI Usage
- AI used: GPT-5.3-Codex (README/script update), Claude Opus 4.6 (historical development assistance)
- Usage scope: documentation restructuring, wording cleanup, language option implementation
- Human review: parameters, command examples, and behavior descriptions were reviewed against `restore_metadata.ps1`

## Contributing
Issues and pull requests are welcome.
Please include:
- Goal of the change
- Summary of implementation
- Reproduction or verification steps

## License
MIT License. See `LICENSE`.

## Support
Open an issue in this repository for bug reports or questions.
