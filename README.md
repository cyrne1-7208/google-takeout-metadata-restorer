# Google Takeout Metadata Restorer

**[日本語版 README はこちら](README_ja.md)**

Restore EXIF metadata (dates, GPS, descriptions) from Google Takeout's supplemental JSON files back into your photos and videos.

When you export via [Google Takeout](https://takeout.google.com/), metadata is stored in separate `.supplemental-metadata.json` files instead of the media files themselves. This tool writes it back using [ExifTool](https://exiftool.org/).

## Features

- **Smart filename matching** — Handles truncated/mangled filenames from Google Takeout
- **Multi-threaded** — Parallel ExifTool execution with `-Threads`
- **Year/Month output** — Organize into `YYYY/MM/` folders with `-OutputPath`
- **Extension auto-fix** — Detects real file type via magic bytes
- **Unicode-safe** — Correctly handles Japanese/CJK filenames
- **WhatIf mode** — Dry-run with full statistics
- **CSV logging** — Per-file results with match methods and error details

## Requirements

- [ExifTool](https://exiftool.org/)
- PowerShell 5.1+ (Windows built-in) or PowerShell 7+

## Installation

### ExifTool

```powershell
# Windows
winget install exiftool

# macOS
brew install exiftool

# Linux
sudo apt install libimage-exiftool-perl
```

<details>
<summary>Windows manual install</summary>

1. Download from [exiftool.org](https://exiftool.org/)
2. Rename `exiftool(-k).exe` → `exiftool.exe`
3. Place in a directory in your `PATH`

</details>

### This script

```powershell
git clone https://github.com/cyrne1-7208/google-takeout-metadata-restorer.git
cd google-takeout-metadata-restorer
```

## Usage

```powershell
# Basic (modifies files in-place)
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos"

# Output to year/month folders (originals untouched)
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos"

# Dry-run preview
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -WhatIf

# Multi-threaded
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos" -Threads 8
```

> If you get a security error: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process`

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-PhotosPath` | *(required)* | Google Takeout photos folder |
| `-OutputPath` | | Output folder (`YYYY/MM/` hierarchy) |
| `-Threads` | `1` | Parallel threads (1–32) |
| `-WhatIf` | | Dry-run mode |
| `-ExifToolPath` | `exiftool` | Path to ExifTool |
| `-NoBackup` | | Skip `_original` backup files |
| `-OriginalFileAction` | `Keep` | Backup handling: `Keep` / `Rename` / `Delete` |
| `-LogFile` | `restore-metadata-log.csv` | CSV log path |
| `-Extensions` | 14 types | Media extensions to process |

## Restored Metadata

| Metadata | EXIF Tags |
|----------|-----------|
| Date taken | `DateTimeOriginal`, `CreateDate` |
| Date modified | `ModifyDate`, `FileModifyDate` |
| GPS | `GPSLatitude/Longitude/Altitude` + Ref |
| Description | `ImageDescription`, `XPComment` |
| Title | `Title`, `XPTitle` |

## Acknowledgements

This project is based on [pfilbin90/google-takeout-metadata-restorer](https://github.com/pfilbin90/google-takeout-metadata-restorer). This fork was created independently under the MIT license without explicit permission from the original author.

Developed with the assistance of [Claude Opus 4.6](https://claude.ai/) (AI by Anthropic).

## License

MIT — See [LICENSE](LICENSE).
