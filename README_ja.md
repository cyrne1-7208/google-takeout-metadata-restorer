# Google Takeout メタデータ復元ツール

**[English README](README.md)**

Google Takeout の `.supplemental-metadata.json` から、写真・動画ファイルに EXIF メタデータ（撮影日時、GPS、説明、タイトル）を書き戻す PowerShell スクリプトです。
他の写真管理アプリへ移行するときに、日時順や位置情報を維持しやすくします。

## 概要
Google Takeout では、メディア本体とメタデータが別ファイルとして出力されます。移行先によっては JSON のメタデータが無視され、撮影日や位置情報が失われます。
このツールは [ExifTool](https://exiftool.org/) を使って、JSON 側の情報をメディアファイルの EXIF タグへ再適用します。

### 対象ユーザー
- Google Takeout から写真・動画を移行するユーザー
- 撮影日時や位置情報を保持したいユーザー

## 特徴
- Google Takeout の切り詰めファイル名に対応した段階的マッチング
- `-Threads`（`1` から `32`）による並列 ExifTool 実行
- `-OutputPath` 指定時の `YYYY/MM/` 階層出力
- マジックバイト判定による拡張子の自動補正
- 日本語/CJK を含む Unicode ファイル名に対応
- `-WhatIf` によるドライラン
- CSV ログと失敗レポートの出力
- `-Language` による表示言語切替（日本語/英語）

## 必要環境
- PowerShell 5.1 以上（Windows）または PowerShell 7+
- [ExifTool](https://exiftool.org/)

## ステータス
- README 最終更新日: 2026-03-09
- 対応環境: PowerShell 5.1 以上 / ExifTool 安定版

## インストール
1. リポジトリを取得します。
2. ExifTool をインストールします。

```powershell
git clone https://github.com/cyrne1-7208/google-takeout-metadata-restorer.git
cd google-takeout-metadata-restorer
```

ExifTool インストール例:

```powershell
# Windows
winget install exiftool

# macOS
brew install exiftool

# Ubuntu/Debian
sudo apt install libimage-exiftool-perl
```

## クイックスタート
まずはファイルを変更しない `-WhatIf` で確認します。

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -WhatIf -Language en
```

期待結果:

```text
マッチング統計、更新予定件数、失敗サマリーが表示されます。
WhatIf では実ファイルは変更されません。
```

## 使い方
通常実行（インプレース更新）:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos"
```

年月フォルダへ出力（元ファイル保持）:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos"
```

並列実行:

```powershell
.\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos" -Threads 8
```

実行ポリシーによりブロックされる場合:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

## 設定
| パラメータ | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `-PhotosPath` | 必須 | なし | Google Takeout のメディアルートフォルダ |
| `-Extensions` | 任意 | 15 種類 | 処理対象の拡張子 |
| `-ExifToolPath` | 任意 | `exiftool` | ExifTool 実行ファイルのパス |
| `-WhatIf` | 任意 | `false` | ドライラン（ファイル変更なし） |
| `-NoBackup` | 任意 | `false` | `_original` バックアップを作成しない |
| `-OriginalFileAction` | 任意 | `Keep` | `_original` の扱い（`Keep` / `Rename` / `Delete`） |
| `-OutputPath` | 任意 | 空 | `YYYY/MM/` 階層での出力先ルート |
| `-LogFile` | 任意 | `restore-metadata-log.csv` | CSV ログ出力先 |
| `-PrefixMatchChars` | 任意 | `20` | プレフィックス一致に使う文字数 |
| `-TimeToleranceSeconds` | 任意 | `86400` | タイムスタンプ近傍判定の許容秒数 |
| `-Threads` | 任意 | `1` | 並列ワーカー数（`1` から `32`） |
| `-Language` | 任意 | `ja` | 表示言語（`ja` / `en`） |

## 復元されるメタデータ
| 項目 | EXIF タグ |
|---|---|
| 撮影日時 | `DateTimeOriginal`, `CreateDate` |
| 更新日時 | `ModifyDate`, `FileModifyDate` |
| GPS | `GPSLatitude`, `GPSLongitude`, `GPSAltitude`（＋Ref） |
| 説明 | `ImageDescription`, `XPComment` |
| タイトル | `Title`, `XPTitle` |

## プロジェクト構成
| パス | 役割 |
|---|---|
| `restore_metadata.ps1` | メタデータ復元の本体スクリプト |
| `README.md` | 英語ドキュメント |
| `README_ja.md` | 日本語ドキュメント |
| `LICENSE` | MIT ライセンス本文 |

## テスト
このリポジトリには自動テストスイートはありません。
小さな Takeout サンプルで `-WhatIf` を先に実行し、結果を確認してから本実行してください。

## トラブルシューティング
- `Execution policy` エラー: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process` を実行します。
- `exiftool` が見つからない: `-ExifToolPath` で `exiftool.exe` のフルパスを指定します。
- 影響範囲が不安: `-WhatIf` で事前確認し、CSV と failure report を確認します。
- 英語表示にしたい: `-Language en` を指定します。

## 謝辞
- ベース実装: [pfilbin90/google-takeout-metadata-restorer](https://github.com/pfilbin90/google-takeout-metadata-restorer)
- メタデータ処理ツール: [ExifTool](https://exiftool.org/)

## 生成支援ツール利用
- 使用 AI: GPT-5.3-Codex（README/スクリプト更新）、Claude Opus 4.6（過去の開発支援）
- 利用範囲: ドキュメント構成整理、文言調整、言語オプション実装
- 人手レビュー: `restore_metadata.ps1` と照合して、パラメータ名・コマンド例・説明を確認

## コントリビュート
Issue / Pull Request を歓迎します。
PR には以下を含めてください。
- 変更目的
- 実装内容の要約
- 再現手順または確認手順

## ライセンス
MIT ライセンスです。詳細は `LICENSE` を参照してください。

## サポート
不具合報告や質問は、このリポジトリの Issue へ投稿してください。
