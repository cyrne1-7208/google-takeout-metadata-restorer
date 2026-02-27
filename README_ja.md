# Google Takeout メタデータ復元ツール

**[English README](README.md)**

Google Takeout でエクスポートした写真・動画の EXIF メタデータ（撮影日時、GPS、説明文）を、付属の JSON ファイルから復元するスクリプトです。

Google Takeout で写真をエクスポートすると、メタデータが `.supplemental-metadata.json` という別ファイルに分離されます。そのまま他のアプリにインポートすると日付がバラバラになり、位置情報も消えます。このツールは [ExifTool](https://exiftool.org/) を使ってメタデータを書き戻します。

## 特徴

- **スマートマッチング** — Google Takeout の切り詰めファイル名にも対応する8段階マッチング
- **マルチスレッド** — `-Threads` で並列 ExifTool 実行
- **年月フォルダ出力** — `-OutputPath` で `YYYY/MM/` 階層に整理（元ファイルは変更なし）
- **拡張子自動修正** — マジックバイトで実際のファイル形式を検出
- **Unicode 対応** — 日本語ファイル名を安全に処理
- **WhatIf モード** — 変更なしのシミュレーション実行
- **CSV ログ** — ファイルごとの処理結果・マッチ方法・エラー詳細を記録

## 必要なもの

- [ExifTool](https://exiftool.org/)
- PowerShell 5.1 以上（Windows 標準搭載）または PowerShell 7+

## インストール

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
<summary>Windows 手動インストール</summary>

1. [exiftool.org](https://exiftool.org/) からダウンロード
2. `exiftool(-k).exe` を `exiftool.exe` にリネーム
3. `PATH` の通ったディレクトリに配置

</details>

### 本スクリプト

```powershell
git clone https://github.com/cyrne1-7208/google-takeout-metadata-restorer.git
cd google-takeout-metadata-restorer
```

## 使い方

```powershell
# 基本（ファイルを直接変更）
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google フォト"

# 年月フォルダに出力（元ファイルはそのまま）
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google フォト" -OutputPath "D:\Photos"

# プレビュー（変更なし）
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google フォト" -WhatIf

# マルチスレッド
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google フォト" -OutputPath "D:\Photos" -Threads 8
```

> 実行ポリシーエラーが出た場合: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process`

## パラメータ

| パラメータ | デフォルト | 説明 |
|-----------|-----------|------|
| `-PhotosPath` | *(必須)* | Google Takeout の写真フォルダ |
| `-OutputPath` | | `YYYY/MM/` 出力先フォルダ |
| `-Threads` | `1` | 並列スレッド数（1〜32） |
| `-WhatIf` | | ドライランモード |
| `-ExifToolPath` | `exiftool` | ExifTool のパス |
| `-NoBackup` | | `_original` バックアップを作成しない |
| `-OriginalFileAction` | `Keep` | バックアップ処理: `Keep` / `Rename` / `Delete` |
| `-LogFile` | `restore-metadata-log.csv` | CSV ログ出力先 |
| `-Extensions` | 14種類 | 処理対象の拡張子 |

## 復元されるメタデータ

| メタデータ | EXIF タグ |
|-----------|-----------|
| 撮影日時 | `DateTimeOriginal`, `CreateDate` |
| 更新日時 | `ModifyDate`, `FileModifyDate` |
| GPS | `GPSLatitude/Longitude/Altitude` + Ref |
| 説明文 | `ImageDescription`, `XPComment` |
| タイトル | `Title`, `XPTitle` |

## 謝辞

本プロジェクトは [pfilbin90/google-takeout-metadata-restorer](https://github.com/pfilbin90/google-takeout-metadata-restorer) をベースにしています。MIT ライセンスに基づき独自にフォークしたものであり、原作者からの明示的な許可は得ていません。

AI である [Claude Opus 4.6](https://claude.ai/)（Anthropic）の支援を受けて開発されました。

## ライセンス

MIT — [LICENSE](LICENSE) を参照。
