<#
.SYNOPSIS
  Google Takeout の JSON メタデータを写真・動画ファイルに復元する。

.DESCRIPTION
  Google Takeout エクスポート時に分離された .supplemental-metadata.json から、
  撮影日時・GPS・説明文等を ExifTool で元のメディアファイルに書き戻します。

  ファイル名マッチング（8 段階）:
    0) JSON ファイル名からメディアファイル名を推定（サフィックス除去）
    1) title フィールド完全一致
    2) Unicode 正規化一致
    3) (N) 重複番号除去一致
    4) ベース名完全一致
    5) プレフィックス一致（同一ディレクトリ優先）
    6) 部分文字列一致
    7) タイムスタンプ近傍（最終手段）

  GPS 0/0 はスキップ（geoDataExif にフォールバック）。

.PARAMETER PhotosPath
  Google Takeout の写真フォルダパス（必須）。

.PARAMETER OutputPath
  YYYY/MM 階層の出力先フォルダ。指定時は元ファイルを変更せずコピー。

.PARAMETER Threads
  並列 ExifTool スレッド数（1〜32、デフォルト: 1）。

.PARAMETER WhatIf
  ドライランモード。ファイルを変更せず統計のみ出力。

.EXAMPLE
  .\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos"

.EXAMPLE
  .\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -OutputPath "D:\Photos" -Threads 8

.EXAMPLE
  .\restore_metadata.ps1 -PhotosPath "C:\Takeout\Google Photos" -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PhotosPath,

    [string[]]$Extensions = @(
        ".jpg", ".jpeg", ".heic", ".png", ".tif", ".tiff",
        ".mp4", ".mov", ".gif", ".webp", ".bmp",
        ".3gp", ".mkv", ".avi", ".avif"
    ),

    [string]$ExifToolPath = "exiftool",

    [switch]$WhatIf,

    [switch]$NoBackup,

    [ValidateSet('Keep','Rename','Delete')]
    [string]$OriginalFileAction = 'Keep',

    [string]$OutputPath = '',

    [string]$LogFile = "$PWD\restore-metadata-log.csv",

    [int]$PrefixMatchChars = 20,

    [int]$TimeToleranceSeconds = 86400,

    [ValidateRange(1, 32)]
    [int]$Threads = 1
)

# ══════════════════════════════════════════════════════════
# 初期設定
# ══════════════════════════════════════════════════════════

$script:UnixEpoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
$startTime = Get-Date

$useOutputPath = $false
if ($OutputPath -and $OutputPath.Trim().Length -gt 0) {
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath.Trim())
    $useOutputPath = $true
    Write-Host "出力先フォルダ: $OutputPath (年/月階層で保存)"
}
$script:UseOutputPath = $useOutputPath

if ($Threads -gt 1) {
    Write-Host "並列処理: $Threads スレッド"
}

# ══════════════════════════════════════════════════════════
# ユーティリティ関数
# ══════════════════════════════════════════════════════════

function Detect-RealExtension {
    <# ファイルのマジックバイトを読み取り、実際のファイル形式に対応する拡張子を返す。
       判定できない場合は $null を返す。 #>
    param([string]$filePath)
    try {
        $fs = [System.IO.File]::OpenRead($filePath)
        try {
            $header = New-Object byte[] 16
            $bytesRead = $fs.Read($header, 0, 16)
            if ($bytesRead -lt 4) { return $null }

            # JPEG: FF D8 FF
            if ($header[0] -eq 0xFF -and $header[1] -eq 0xD8 -and $header[2] -eq 0xFF) { return '.jpg' }
            # PNG: 89 50 4E 47
            if ($header[0] -eq 0x89 -and $header[1] -eq 0x50 -and $header[2] -eq 0x4E -and $header[3] -eq 0x47) { return '.png' }
            # GIF: 47 49 46 38
            if ($header[0] -eq 0x47 -and $header[1] -eq 0x49 -and $header[2] -eq 0x46 -and $header[3] -eq 0x38) { return '.gif' }
            # BMP: 42 4D
            if ($header[0] -eq 0x42 -and $header[1] -eq 0x4D) { return '.bmp' }
            # TIFF: II (49 49 2A 00) or MM (4D 4D 00 2A)
            if (($header[0] -eq 0x49 -and $header[1] -eq 0x49 -and $header[2] -eq 0x2A -and $header[3] -eq 0x00) -or
                ($header[0] -eq 0x4D -and $header[1] -eq 0x4D -and $header[2] -eq 0x00 -and $header[3] -eq 0x2A)) { return '.tif' }
            # WEBP: RIFF....WEBP
            if ($bytesRead -ge 12 -and
                $header[0] -eq 0x52 -and $header[1] -eq 0x49 -and $header[2] -eq 0x46 -and $header[3] -eq 0x46 -and
                $header[8] -eq 0x57 -and $header[9] -eq 0x45 -and $header[10] -eq 0x42 -and $header[11] -eq 0x50) { return '.webp' }
            # HEIC/HEIF/MP4/MOV/3GP: ftyp box
            if ($bytesRead -ge 12 -and
                $header[4] -eq 0x66 -and $header[5] -eq 0x74 -and $header[6] -eq 0x79 -and $header[7] -eq 0x70) {
                $brand = [System.Text.Encoding]::ASCII.GetString($header, 8, 4)
                if ($brand -match 'avif|avis') { return '.avif' }
                if ($brand -match 'heic|heix|mif1') { return '.heic' }
                if ($brand -match 'isom|mp4|avc1|M4V|dash') { return '.mp4' }
                if ($brand -match 'qt|M4A') { return '.mov' }
                if ($brand -match '3gp|3g2') { return '.3gp' }
                return '.mp4'
            }
            # MKV: 1A 45 DF A3
            if ($header[0] -eq 0x1A -and $header[1] -eq 0x45 -and $header[2] -eq 0xDF -and $header[3] -eq 0xA3) { return '.mkv' }
            # AVI: RIFF....AVI
            if ($bytesRead -ge 12 -and
                $header[0] -eq 0x52 -and $header[1] -eq 0x49 -and $header[2] -eq 0x46 -and $header[3] -eq 0x46 -and
                $header[8] -eq 0x41 -and $header[9] -eq 0x56 -and $header[10] -eq 0x49 -and $header[11] -eq 0x20) { return '.avi' }

            return $null
        }
        finally { $fs.Close() }
    }
    catch { return $null }
}

function Normalize-String {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    try   { return $s.Normalize([System.Text.NormalizationForm]::FormC) }
    catch { return $s }
}

function Escape-CsvField {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '""' }
    return '"' + $s.Replace('"', '""') + '"'
}

function Parse-JsonFile {
    param([string]$jsonPath)
    try {
        $text = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 -ErrorAction Stop
        return ConvertFrom-Json -InputObject $text -ErrorAction Stop
    }
    catch { return $null }
}

function To-ExifDateTime {
    param([long]$unixTimestamp)
    if ($unixTimestamp -eq 0) { return $null }
    try {
        $dt = $script:UnixEpoch.AddSeconds([double]$unixTimestamp).ToLocalTime()
        return $dt.ToString('yyyy:MM:dd HH:mm:ss')
    }
    catch { return $null }
}

function Test-GooglePhotosJson {
    <# JSON が Google Photos メタデータ形式か簡易判定 #>
    param([string]$path)
    try {
        $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
        return ($text -match '"photoTakenTime"' -or $text -match '"creationTime"')
    }
    catch { return $false }
}

# ══════════════════════════════════════════════════════════
# ファイル名パターンマッチング関数
# ══════════════════════════════════════════════════════════

function Get-MediaNameFromJsonName {
    <#
    JSONファイル名からメディアファイル名を推定する。
    Google Takeoutの切り詰めパターンに対応:
      正常:    photo.jpg.supplemental-metadata.json       → photo.jpg
      (N)付:   cachedImage.png.supplemental-metadata(1).json → cachedImage.png (DupNum=1)
      切詰1:   photo.jpg.suppl.json / photo.jpg.supp.json → photo.jpg
      切詰2:   photo.jpg.sup.json / photo.jpg.s.json      → photo.jpg
      切詰3:   photo.jpg..json (サフィックス完全消失)       → photo.jpg
      切詰4:   photo.jpg.supplementa.json                 → photo.jpg
      切詰5:   photo.supp(1).json                         → photo (DupNum=1)
      極端:    photo.json (拡張子ごと消失)                  → photo (IsTruncated)
    #>
    param([string]$jsonFileName)
    $name = $jsonFileName

    # ── 1) supplemental-metadata(N).json ──
    if ($name -match '^(.+)\.supplemental-metadata\((\d+)\)\.json$') {
        return @{ Name = $Matches[1]; DupNum = $Matches[2]; IsTruncated = $false }
    }

    # ── 2) 既知サフィックスの完全一致除去（長い順） ──
    $suffixes = @('.supplemental-metadata.json', '.suppl.json', '.supp.json')
    foreach ($sfx in $suffixes) {
        if ($name.EndsWith($sfx, [StringComparison]::OrdinalIgnoreCase)) {
            return @{ Name = $name.Substring(0, $name.Length - $sfx.Length); DupNum = $null; IsTruncated = $false }
        }
    }

    # ── 3) 汎用 .supplementalXXX.json（切り詰め含む）──
    if ($name -match '^(.+)\.supplement[a-z]*\.json$') {
        return @{ Name = $Matches[1]; DupNum = $null; IsTruncated = $false }
    }

    # ── 4) .supp(N).json / .suppl(N).json パターン ──
    if ($name -match '^(.+)\.supp[l]?\((\d+)\)\.json$') {
        return @{ Name = $Matches[1]; DupNum = $Matches[2]; IsTruncated = $false }
    }

    # ── 5) 切り詰めサフィックス: .sup.json, .su.json, .s.json ──
    if ($name -match '^(.+)\.(sup|su|s)\.json$') {
        return @{ Name = $Matches[1]; DupNum = $null; IsTruncated = $false }
    }

    # ── 6) ドット二重: photo.jpg..json ──
    if ($name -match '^(.+)\.\.json$') {
        return @{ Name = $Matches[1]; DupNum = $null; IsTruncated = $false }
    }

    # ── 7) (N).json パターン ──
    if ($name -match '^(.+)\.\((\d+)\)\.json$') {
        return @{ Name = $Matches[1]; DupNum = $Matches[2]; IsTruncated = $false }
    }

    # ── 8) メディア拡張子.json ──
    $mediaExtPattern = '\.(jpg|jpeg|heic|png|tif|tiff|mp4|mov|gif|webp|bmp|3gp|mkv|avi)\.json$'
    if ($name -match $mediaExtPattern) {
        return @{ Name = $name.Substring(0, $name.Length - 5); DupNum = $null; IsTruncated = $false }
    }

    # ── 9) 最終フォールバック：.json のみ除去 ──
    if ($name.EndsWith('.json', [StringComparison]::OrdinalIgnoreCase)) {
        return @{ Name = $name.Substring(0, $name.Length - 5); DupNum = $null; IsTruncated = $true }
    }

    return @{ Name = $name; DupNum = $null; IsTruncated = $false }
}

function Remove-DuplicateSuffix {
    <# ファイル名から (N) 重複サフィックスを除去する #>
    param([string]$filename)
    if ($filename -match '^(.+?)\(\d+\)(\.[^.]+)$') { return $Matches[1] + $Matches[2] }
    if ($filename -match '^(.+?)\(\d+\)$') { return $Matches[1] }
    return $filename
}

# ══════════════════════════════════════════════════════════
# メディアファイル検索関数
# ══════════════════════════════════════════════════════════

function Find-MediaCandidate {
    param(
        [string]$titleFilename,
        [string]$jsonPath,
        [hashtable]$filesByName,
        [hashtable]$filesByBase,
        [hashtable]$filesByDir,
        [ref]$matchMethod
    )

    $title         = $titleFilename
    $titleNorm     = (Normalize-String $title).ToLowerInvariant()
    $baseTitle     = [System.IO.Path]::GetFileNameWithoutExtension($title)
    $baseTitleNorm = (Normalize-String $baseTitle).ToLowerInvariant()
    $jsonDir       = (Split-Path -Parent $jsonPath).ToLowerInvariant()

    # ── 同一ディレクトリ優先の選択ヘルパー ──
    function Select-PreferSameDir {
        param([string[]]$paths)
        if ($paths.Count -eq 0) { return $null }
        if ($paths.Count -eq 1) { return $paths[0] }
        foreach ($p in $paths) {
            if ((Split-Path -Parent $p).ToLowerInvariant() -eq $jsonDir) { return $p }
        }
        return $paths[0]
    }

    # ── 0) JSON ファイル名からメディア名を推定 ──
    $derived    = Get-MediaNameFromJsonName -jsonFileName (Split-Path -Leaf $jsonPath)
    $derivedName   = $derived.Name
    $derivedDupNum = $derived.DupNum

    if ($derivedName -and $derivedName.Length -gt 0) {
        # 0a) supplemental-metadata(N).json → メディア名に(N)付与
        if ($derivedDupNum) {
            $ext  = [System.IO.Path]::GetExtension($derivedName)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($derivedName)
            $dupMediaLower = "${base}(${derivedDupNum})${ext}".ToLowerInvariant()
            if ($filesByName.ContainsKey($dupMediaLower)) {
                $matchMethod.Value = 'JsonFilename-DupNum'
                return (Select-PreferSameDir $filesByName[$dupMediaLower])
            }
        }

        # 0b) 推定名で完全一致
        $derivedLower = $derivedName.ToLowerInvariant()
        if ($filesByName.ContainsKey($derivedLower)) {
            $matchMethod.Value = 'JsonFilename-Exact'
            return (Select-PreferSameDir $filesByName[$derivedLower])
        }

        # 0c) (N) 除去して再試行
        $derivedClean      = Remove-DuplicateSuffix $derivedName
        $derivedCleanLower = $derivedClean.ToLowerInvariant()
        if ($derivedCleanLower -ne $derivedLower -and $filesByName.ContainsKey($derivedCleanLower)) {
            $matchMethod.Value = 'JsonFilename-DedupSuffix'
            return (Select-PreferSameDir $filesByName[$derivedCleanLower])
        }

        # 0d) 拡張子なし/切り詰め → プレフィックスマッチ
        $derivedExt = [System.IO.Path]::GetExtension($derivedName)
        if ([string]::IsNullOrEmpty($derivedExt) -or $derived.IsTruncated) {
            $derivedNorm = (Normalize-String $derivedName).ToLowerInvariant()
            $prefLen = $derivedNorm.Length
            if ($prefLen -ge 4) {
                $pref = $derivedNorm.Substring(0, $prefLen)
                $candidates = [System.Collections.Generic.List[string]]::new()
                foreach ($kv in $filesByName.GetEnumerator()) {
                    if ($kv.Key.StartsWith($pref)) {
                        foreach ($p in $kv.Value) { $candidates.Add($p) }
                    }
                }
                $sameDirCands = @($candidates | Where-Object {
                    (Split-Path -Parent $_).ToLowerInvariant() -eq $jsonDir
                })
                if ($sameDirCands.Count -eq 1) {
                    $matchMethod.Value = 'JsonFilename-TruncPrefix-SameDir'
                    return $sameDirCands[0]
                }
                if ($candidates.Count -eq 1) {
                    $matchMethod.Value = 'JsonFilename-TruncPrefix-Global'
                    return $candidates[0]
                }
            }
        }
    }

    # ── 1) title 完全一致（case-insensitive） ──
    $lower = $title.ToLowerInvariant()
    if ($filesByName.ContainsKey($lower)) {
        $matchMethod.Value = 'Title-Exact'
        return (Select-PreferSameDir $filesByName[$lower])
    }

    # ── 2) Unicode 正規化で一致 ──
    foreach ($k in $filesByName.Keys) {
        if ((Normalize-String $k) -eq $titleNorm) {
            $matchMethod.Value = 'Title-Normalized'
            return (Select-PreferSameDir $filesByName[$k])
        }
    }

    # ── 3) (N) 重複番号除去で一致 ──
    $titleClean      = Remove-DuplicateSuffix $title
    $titleCleanLower = $titleClean.ToLowerInvariant()
    if ($titleCleanLower -ne $lower) {
        if ($filesByName.ContainsKey($titleCleanLower)) {
            $matchMethod.Value = 'Title-DedupSuffix'
            return (Select-PreferSameDir $filesByName[$titleCleanLower])
        }
        # 逆方向: メディア側の (N) を除去して一致確認
        foreach ($k in $filesByName.Keys) {
            $kClean = (Remove-DuplicateSuffix $k).ToLowerInvariant()
            if ($kClean -eq $lower -and $kClean -ne $k) {
                $matchMethod.Value = 'Title-ReverseDedupSuffix'
                return (Select-PreferSameDir $filesByName[$k])
            }
        }
    }

    # ── 4) ベース名完全一致 ──
    foreach ($kv in $filesByBase.GetEnumerator()) {
        if ((Normalize-String $kv.Key) -eq $baseTitleNorm) {
            $matchMethod.Value = 'Base-Exact'
            return (Select-PreferSameDir $kv.Value)
        }
    }

    # ── 5) プレフィックス一致 ──
    $prefixLen = [Math]::Min($baseTitleNorm.Length, $PrefixMatchChars)
    if ($prefixLen -ge 5) {
        $prefix = $baseTitleNorm.Substring(0, $prefixLen)
        $candidates = [System.Collections.Generic.List[string]]::new()
        foreach ($kv in $filesByBase.GetEnumerator()) {
            if ($kv.Key.StartsWith($prefix)) {
                foreach ($p in $kv.Value) { $candidates.Add($p) }
            }
        }
        $sameDirCandidates = @($candidates | Where-Object {
            (Split-Path -Parent $_).ToLowerInvariant() -eq $jsonDir
        })
        if ($sameDirCandidates.Count -eq 1) {
            $matchMethod.Value = 'Prefix-SameDir'
            return $sameDirCandidates[0]
        }
        if ($candidates.Count -eq 1) {
            $matchMethod.Value = 'Prefix-Global'
            return $candidates[0]
        }
    }

    # ── 6) 部分文字列一致（先頭12文字を含む） ──
    $subLen = [Math]::Min($baseTitleNorm.Length, 12)
    if ($subLen -ge 5) {
        $sub = $baseTitleNorm.Substring(0, $subLen)
        $candidates = [System.Collections.Generic.List[string]]::new()
        foreach ($kv in $filesByBase.GetEnumerator()) {
            if ($kv.Key.Contains($sub)) {
                foreach ($p in $kv.Value) { $candidates.Add($p) }
            }
        }
        $sameDirCandidates = @($candidates | Where-Object {
            (Split-Path -Parent $_).ToLowerInvariant() -eq $jsonDir
        })
        if ($sameDirCandidates.Count -eq 1) {
            $matchMethod.Value = 'Contains-SameDir'
            return $sameDirCandidates[0]
        }
        if ($candidates.Count -eq 1) {
            $matchMethod.Value = 'Contains-Global'
            return $candidates[0]
        }
    }

    # ── 7) タイムスタンプ近傍（同一ディレクトリ限定、最終手段） ──
    $jsonObj = Parse-JsonFile -jsonPath $jsonPath
    if ($null -ne $jsonObj -and $jsonObj.photoTakenTime -and $jsonObj.photoTakenTime.timestamp) {
        try {
            $targetTs = [double]$jsonObj.photoTakenTime.timestamp
            $target   = $script:UnixEpoch.AddSeconds($targetTs).ToLocalTime()
        }
        catch { $target = $null }

        if ($target) {
            $pool = @()
            if ($filesByDir.ContainsKey($jsonDir)) { $pool = $filesByDir[$jsonDir] }

            $best = $null; $bestDiff = [double]::MaxValue
            foreach ($fi in $pool) {
                $diff = [Math]::Abs(($fi.LastWriteTime - $target).TotalSeconds)
                if ($diff -lt $bestDiff) { $best = $fi; $bestDiff = $diff }
            }
            if ($best -and $bestDiff -lt $TimeToleranceSeconds) {
                $matchMethod.Value = "Timestamp($([Math]::Round($bestDiff,1))s)"
                return $best.FullName
            }
        }
    }

    $matchMethod.Value = 'NoMatch'
    return $null
}

# ══════════════════════════════════════════════════════════
# 出力・ロギング関数
# ══════════════════════════════════════════════════════════

function Get-OutputDestination {
    <# photoTakenTime / creationTime から YYYY/MM サブフォルダのパスを返す #>
    param(
        [string]$outputRoot,
        [string]$mediaFileName,
        $jsonObj
    )

    # タイムスタンプ取得: photoTakenTime 優先 → creationTime フォールバック
    $ts = $null
    if ($jsonObj.photoTakenTime -and $jsonObj.photoTakenTime.timestamp) {
        try { $ts = [long]$jsonObj.photoTakenTime.timestamp } catch {}
    }
    if ((-not $ts -or $ts -eq 0) -and $jsonObj.creationTime -and $jsonObj.creationTime.timestamp) {
        try { $ts = [long]$jsonObj.creationTime.timestamp } catch {}
    }

    if ($ts -and $ts -ne 0) {
        $dt = $script:UnixEpoch.AddSeconds([double]$ts).ToLocalTime()
        $yearFolder  = $dt.ToString('yyyy')
        $monthFolder = $dt.ToString('MM')
    }
    else {
        $yearFolder  = 'unknown'
        $monthFolder = '00'
    }

    $destDir = Join-Path (Join-Path $outputRoot $yearFolder) $monthFolder
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    # ファイル名の重複回避（ディスク上の既存ファイル＋Phase 2 で割り当て済みパスの両方をチェック）
    $destFile = Join-Path $destDir $mediaFileName
    $destLower = $destFile.ToLowerInvariant()
    if ((Test-Path -LiteralPath $destFile) -or $script:assignedPaths.ContainsKey($destLower)) {
        $ext  = [System.IO.Path]::GetExtension($mediaFileName)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($mediaFileName)
        $counter = 1
        do {
            $destFile = Join-Path $destDir "${base}(${counter})${ext}"
            $destLower = $destFile.ToLowerInvariant()
            $counter++
        } while ((Test-Path -LiteralPath $destFile) -or $script:assignedPaths.ContainsKey($destLower))
    }
    $script:assignedPaths[$destLower] = $true

    return @{ DestDir = $destDir; DestPath = $destFile; Year = $yearFolder; Month = $monthFolder }
}

function New-LogFields {
    <# CSV ログ行のフィールド配列を構築する（OutputPath モード自動判定） #>
    param(
        [string]$JsonPath    = '',
        [string]$MediaPath   = '',
        [string]$MatchMethod = '',
        [string]$Result      = '',
        [string]$PhotoTime   = '',
        [string]$CreateTime  = '',
        [string]$HasGps      = '',
        [string]$Description = '',
        [string]$DestPath    = '',
        [string]$Notes       = ''
    )
    if ($script:UseOutputPath) {
        return @($JsonPath, $MediaPath, $MatchMethod, $Result, $PhotoTime, $CreateTime, $HasGps, $Description, $DestPath, $Notes)
    }
    return @($JsonPath, $MediaPath, $MatchMethod, $Result, $PhotoTime, $CreateTime, $HasGps, $Description, $Notes)
}

function Write-LogLine {
    param([string[]]$Fields)
    $line = ($Fields | ForEach-Object { Escape-CsvField $_ }) -join ','
    Add-Content -Path $LogFile -Value $line
}

# ══════════════════════════════════════════════════════════
# ExifTool 引数構築関数
# ══════════════════════════════════════════════════════════

function Build-ExifToolArgs {
    <# JSON オブジェクトから ExifTool に渡すタグ引数リストを構築する。
       戻り値: @{ Args=[List[string]]; PhotoTime; CreateTime; HasGps; Description } #>
    param($json)

    $tagArgs = [System.Collections.Generic.List[string]]::new()

    # ── タイムスタンプ ──
    # photoTakenTime → 撮影日時 (DateTimeOriginal, CreateDate)
    # creationTime   → 更新日時 (ModifyDate, FileModifyDate)
    $dtPhoto = $null
    if ($json.photoTakenTime -and $json.photoTakenTime.timestamp) {
        $dtPhoto = To-ExifDateTime -unixTimestamp ([long]$json.photoTakenTime.timestamp)
    }
    $dtCreate = $null
    if ($json.creationTime -and $json.creationTime.timestamp) {
        $dtCreate = To-ExifDateTime -unixTimestamp ([long]$json.creationTime.timestamp)
    }

    if ($dtPhoto) {
        $tagArgs.Add("-DateTimeOriginal=$dtPhoto")
        $tagArgs.Add("-CreateDate=$dtPhoto")
    }
    if ($dtCreate) {
        $tagArgs.Add("-ModifyDate=$dtCreate")
        $tagArgs.Add("-FileModifyDate=$dtCreate")
    }
    elseif ($dtPhoto) {
        $tagArgs.Add("-ModifyDate=$dtPhoto")
    }

    # ── 説明・タイトル ──
    $descText = ''
    if ($json.description -and $json.description.Trim().Length -gt 0) {
        $descText = $json.description.Trim()
        $tagArgs.Add("-ImageDescription=$descText")
        $tagArgs.Add("-XPComment=$descText")
    }
    if ($json.title -and $json.title.Trim().Length -gt 0) {
        $t = $json.title.Trim()
        $tagArgs.Add("-Title=$t")
        $tagArgs.Add("-XPTitle=$t")
    }

    # ── GPS（geoData → geoDataExif のフォールバック） ──
    $hasGps    = $false
    $geoSource = $null
    if ($json.geoData) {
        $gLat = 0.0; $gLon = 0.0
        try { $gLat = [double]$json.geoData.latitude }  catch {}
        try { $gLon = [double]$json.geoData.longitude } catch {}
        if ($gLat -ne 0.0 -or $gLon -ne 0.0) { $geoSource = $json.geoData }
    }
    if (-not $geoSource -and $json.geoDataExif) {
        $gLat = 0.0; $gLon = 0.0
        try { $gLat = [double]$json.geoDataExif.latitude }  catch {}
        try { $gLon = [double]$json.geoDataExif.longitude } catch {}
        if ($gLat -ne 0.0 -or $gLon -ne 0.0) { $geoSource = $json.geoDataExif }
    }
    if ($geoSource) {
        $hasGps = $true
        $lat = [double]$geoSource.latitude
        $lon = [double]$geoSource.longitude
        $alt = 0.0
        try { $alt = [double]$geoSource.altitude } catch {}

        $latRef = if ($lat -ge 0) { 'N' } else { 'S' }
        $lonRef = if ($lon -ge 0) { 'E' } else { 'W' }

        $tagArgs.Add("-GPSLatitude=$([Math]::Abs($lat))")
        $tagArgs.Add("-GPSLatitudeRef=$latRef")
        $tagArgs.Add("-GPSLongitude=$([Math]::Abs($lon))")
        $tagArgs.Add("-GPSLongitudeRef=$lonRef")

        if ($alt -ne 0.0) {
            $altRef = if ($alt -ge 0) { '0' } else { '1' }
            $tagArgs.Add("-GPSAltitude=$([Math]::Abs($alt))")
            $tagArgs.Add("-GPSAltitudeRef=$altRef")
        }
    }

    return @{
        Args        = $tagArgs
        PhotoTime   = if ($dtPhoto)  { $dtPhoto }  else { '' }
        CreateTime  = if ($dtCreate) { $dtCreate } else { '' }
        HasGps      = if ($hasGps)   { 'Yes' }     else { 'No' }
        Description = $descText
    }
}

# ══════════════════════════════════════════════════════════
# Phase 1: スキャン（メディア・JSONファイルの収集）
# ══════════════════════════════════════════════════════════

if (-not (Test-Path -LiteralPath $PhotosPath)) {
    throw "PhotosPath not found: $PhotosPath"
}

# ── ログ初期化 ──
$logDir = Split-Path -Parent $LogFile
if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$csvHeader = if ($useOutputPath) {
    'JsonPath,MediaPath,MatchMethod,Result,PhotoTakenTime,CreationTime,HasGPS,Description,DestPath,Notes'
} else {
    'JsonPath,MediaPath,MatchMethod,Result,PhotoTakenTime,CreationTime,HasGPS,Description,Notes'
}
Set-Content -Path $LogFile -Value $csvHeader -Encoding UTF8

# ── メディアファイルの収集 ──
Write-Host "Phase 1: ファイルスキャン..."
Write-Host "  Scanning media files under: $PhotosPath"
$allFiles = @(Get-ChildItem -LiteralPath $PhotosPath -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $Extensions -contains $_.Extension.ToLowerInvariant() })
Write-Host "  Found $($allFiles.Count) media files."

# ── ヘルパーマップの構築 ──
$filesByName = @{}   # filename.ext (lower) → @(FullPath, ...)
$filesByBase = @{}   # basename     (lower) → @(FullPath, ...)
$filesByDir  = @{}   # dirpath      (lower) → @(FileInfo, ...)

foreach ($f in $allFiles) {
    $nameLower = $f.Name.ToLowerInvariant()
    if (-not $filesByName.ContainsKey($nameLower)) { $filesByName[$nameLower] = @() }
    $filesByName[$nameLower] += $f.FullName

    $baseLower = $f.BaseName.ToLowerInvariant()
    if (-not $filesByBase.ContainsKey($baseLower)) { $filesByBase[$baseLower] = @() }
    $filesByBase[$baseLower] += $f.FullName

    $dirLower = $f.DirectoryName.ToLowerInvariant()
    if (-not $filesByDir.ContainsKey($dirLower)) { $filesByDir[$dirLower] = @() }
    $filesByDir[$dirLower] += $f   # FileInfo を保持（タイムスタンプ比較用）
}

# ── JSON ファイルの検索・検証 ──
Write-Host "  Finding metadata JSON files..."
$allJsonCandidates = @(Get-ChildItem -LiteralPath $PhotosPath -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -eq '.json' })
Write-Host "  Found $($allJsonCandidates.Count) JSON files total. Validating content..."

$jsonFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$validatedCount = 0
foreach ($jc in $allJsonCandidates) {
    $validatedCount++
    if ($validatedCount % 200 -eq 0) {
        Write-Host "  Validating... $validatedCount / $($allJsonCandidates.Count)"
    }
    if (Test-GooglePhotosJson -path $jc.FullName) {
        $jsonFiles.Add($jc)
    }
}

if ($jsonFiles.Count -eq 0) {
    Write-Host "No Google Photos metadata JSON files found."
    exit
}
Write-Host "  Validated: $($jsonFiles.Count) Google Photos metadata JSON files."

# ══════════════════════════════════════════════════════════
# Phase 2: マッチング＆作業アイテム構築
# ══════════════════════════════════════════════════════════

# 出力先パスの重複回避用（Phase 2 で割り当て済みパスを追跡）
$script:assignedPaths = @{}

Write-Host ""
Write-Host "Phase 2: JSON とメディアファイルのマッチング..."

$workItems     = [System.Collections.Generic.List[hashtable]]::new()
$failJsonParse = [System.Collections.Generic.List[string]]::new()
$failNoMedia   = [System.Collections.Generic.List[psobject]]::new()
$failNoMeta    = [System.Collections.Generic.List[string]]::new()

for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
    $jf = $jsonFiles[$i]

    # ── プログレス ──
    if (($i + 1) % 100 -eq 0 -or $i -eq 0 -or $i -eq $jsonFiles.Count - 1) {
        Write-Progress -Activity "マッチング" `
            -Status "$($i + 1) / $($jsonFiles.Count): $($jf.Name)" `
            -PercentComplete ([int](100 * ($i + 1) / $jsonFiles.Count))
    }

    # ── JSON 解析 ──
    $json = Parse-JsonFile -jsonPath $jf.FullName
    if ($null -eq $json) {
        Write-LogLine (New-LogFields -JsonPath $jf.FullName -Result 'JSONParseError' -Notes 'JSON parse failed')
        $failJsonParse.Add($jf.FullName)
        continue
    }

    # ── タイトル（メディアファイル名）の決定 ──
    $title = $null
    if ($json.title -and $json.title.Trim().Length -gt 0) {
        $title = $json.title.Trim()
    }
    elseif ($json.fileName -and $json.fileName.Trim().Length -gt 0) {
        $title = $json.fileName.Trim()
    }
    else {
        $derived = Get-MediaNameFromJsonName -jsonFileName $jf.Name
        $title = $derived.Name
    }

    # ── メディア候補の検索 ──
    [string]$matchMethod = ''
    $candidate = Find-MediaCandidate `
        -titleFilename $title `
        -jsonPath      $jf.FullName `
        -filesByName   $filesByName `
        -filesByBase   $filesByBase `
        -filesByDir    $filesByDir `
        -matchMethod   ([ref]$matchMethod)

    if (-not $candidate -or -not (Test-Path -LiteralPath $candidate)) {
        $reason = if (-not $candidate) { "No matching media file (title=$title)" } else { 'Matched file no longer exists' }
        $failNoMedia.Add([pscustomobject]@{
            JsonPath       = $jf.FullName
            Title          = $title
            MatchAttempted = if ($candidate) { "FileGone:$matchMethod" } else { $matchMethod }
        })
        Write-LogLine (New-LogFields -JsonPath $jf.FullName `
            -MediaPath $(if ($candidate) { $candidate } else { '' }) `
            -MatchMethod $matchMethod -Result 'Skipped' -Notes $reason)
        continue
    }

    # ── ExifTool 引数の構築 ──
    $meta = Build-ExifToolArgs -json $json

    if ($meta.Args.Count -eq 0) {
        $failNoMeta.Add($jf.FullName)
        Write-LogLine (New-LogFields -JsonPath $jf.FullName -MediaPath $candidate `
            -MatchMethod $matchMethod -Result 'Skipped' `
            -PhotoTime $meta.PhotoTime -CreateTime $meta.CreateTime `
            -HasGps $meta.HasGps -Description $meta.Description `
            -Notes 'No metadata to write')
        continue
    }

    # ── 出力先パスの決定 ──
    $destPath = ''
    $extFixed = $false
    $extNote  = ''
    if ($useOutputPath) {
        # マジックバイトで実際のファイル形式を検出し、拡張子不一致なら修正
        $mediaLeaf = Split-Path -Leaf $candidate
        $realExt = Detect-RealExtension -filePath $candidate
        if ($realExt) {
            $currentExt = [System.IO.Path]::GetExtension($mediaLeaf).ToLowerInvariant()
            $normalizeMap = @{ '.jpeg'='.jpg'; '.tiff'='.tif' }
            $normCurrent = if ($normalizeMap.ContainsKey($currentExt)) { $normalizeMap[$currentExt] } else { $currentExt }
            $normReal    = if ($normalizeMap.ContainsKey($realExt))    { $normalizeMap[$realExt] }    else { $realExt }
            if ($normCurrent -ne $normReal) {
                $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($mediaLeaf)
                $mediaLeaf = $baseName + $realExt
                $extFixed  = $true
                $extNote   = "ExtFixed:$currentExt->$realExt"
            }
        }
        $outInfo  = Get-OutputDestination -outputRoot $OutputPath -mediaFileName $mediaLeaf -jsonObj $json
        $destPath = $outInfo.DestPath
    }

    # ── コマンド引数文字列の構築（表示・ログ用） ──
    $cmdParts = [System.Collections.Generic.List[string]]::new()
    if ($NoBackup -or $useOutputPath) { $cmdParts.Add('-overwrite_original') }
    $cmdParts.Add('-P'); $cmdParts.Add('-charset'); $cmdParts.Add('filename=utf8')
    foreach ($a in $meta.Args) { $cmdParts.Add($a) }
    $targetDisplay = if ($useOutputPath) { $destPath } else { $candidate }
    $cmdParts.Add($targetDisplay)
    $argString = ($cmdParts | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '

    # ── 作業アイテムの登録 ──
    $workItems.Add(@{
        Index        = $workItems.Count
        JsonPath     = $jf.FullName
        JsonName     = $jf.Name
        MediaPath    = $candidate
        MatchMethod  = $matchMethod
        ExifArgs     = [string[]]$meta.Args.ToArray()
        DestPath     = $destPath
        ExtFixed     = $extFixed
        ExtNote      = $extNote
        LogPhotoTime = $meta.PhotoTime
        LogCreateTime= $meta.CreateTime
        LogHasGps    = $meta.HasGps
        LogDesc      = $meta.Description
        ArgString    = $argString
    })
}

Write-Progress -Activity "マッチング" -Status "完了" -PercentComplete 100 -Completed
Write-Host ("  マッチ完了: {0} 件の処理対象, {1} JSON解析エラー, {2} メディア不一致, {3} メタデータなし" `
    -f $workItems.Count, $failJsonParse.Count, $failNoMedia.Count, $failNoMeta.Count)

# ══════════════════════════════════════════════════════════
# Phase 3: 実行（ExifTool 書き込み）
# ══════════════════════════════════════════════════════════

$failExifTool = [System.Collections.Generic.List[psobject]]::new()
$successList  = [System.Collections.Generic.List[psobject]]::new()
$totalWork    = $workItems.Count

if ($totalWork -eq 0) {
    Write-Host "`n処理対象のファイルがありません。"
}
elseif ($WhatIf) {
    # ═══════ WhatIf モード ═══════
    Write-Host "`nPhase 3: WhatIf モード（$totalWork 件）..."
    foreach ($item in $workItems) {
        $successList.Add([pscustomobject]@{
            JsonPath    = $item.JsonPath
            MediaPath   = $item.MediaPath
            MatchMethod = $item.MatchMethod
            DestPath    = $item.DestPath
            ExtFixed    = $item.ExtFixed
            Args        = $item.ArgString
        })
        $notes = @("Args: $($item.ArgString)")
        if ($item.ExtNote) { $notes = @($item.ExtNote) + $notes }
        Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
            -MatchMethod $item.MatchMethod -Result 'WhatIf' `
            -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
            -HasGps $item.LogHasGps -Description $item.LogDesc `
            -DestPath $item.DestPath -Notes ($notes -join '; '))
    }
}
elseif ($Threads -le 1) {
    # ═══════ シーケンシャル実行 ═══════
    Write-Host "`nPhase 3: シーケンシャル実行（$totalWork 件）..."
    $completedCount = 0

    foreach ($item in $workItems) {
        $completedCount++
        Write-Progress -Activity "ExifTool 実行中" `
            -Status "$completedCount / $totalWork" `
            -PercentComplete ([int](100 * $completedCount / $totalWork))

        $targetFile  = $item.MediaPath
        $argFilePath = $null

        try {
            # ── OutputPath へのコピー ──
            if ($useOutputPath) {
                try {
                    Copy-Item -LiteralPath $item.MediaPath -Destination $item.DestPath -Force -ErrorAction Stop
                    $targetFile = $item.DestPath
                }
                catch {
                    $failExifTool.Add([pscustomobject]@{
                        JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                        ExitCode = -1; Error = "Copy failed: $($_.Exception.Message)"
                    })
                    Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                        -MatchMethod $item.MatchMethod -Result 'Error' `
                        -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                        -HasGps $item.LogHasGps -Description $item.LogDesc `
                        -DestPath $item.DestPath -Notes "Copy failed: $($_.Exception.Message)")
                    continue
                }
            }

            # ── ExifTool 引数ファイルの構築 ──
            $cmdArgs = [System.Collections.Generic.List[string]]::new()
            if ($NoBackup -or $useOutputPath) { $cmdArgs.Add('-overwrite_original') }
            $cmdArgs.Add('-P')
            $cmdArgs.Add('-charset')
            $cmdArgs.Add('filename=utf8')
            foreach ($a in $item.ExifArgs) { $cmdArgs.Add($a) }
            $cmdArgs.Add($targetFile)

            $argFilePath = [System.IO.Path]::GetTempFileName()
            $utf8NoBom   = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllLines($argFilePath, $cmdArgs.ToArray(), $utf8NoBom)

            # ── ExifTool 実行（argfile 方式） ──
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName               = $ExifToolPath
            $startInfo.Arguments              = "-@ `"$argFilePath`""
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError  = $true
            $startInfo.UseShellExecute        = $false
            $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $startInfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

            $p      = [System.Diagnostics.Process]::Start($startInfo)
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $p.WaitForExit()

            # Warning のみでファイルが更新されている場合は成功扱い
            $isWarningOnly = ($p.ExitCode -ne 0 -and $stderr -match 'Warning:' -and
                              $stderr -notmatch 'Error:' -and $stdout -match '1 image files updated')

            if ($p.ExitCode -eq 0 -or $isWarningOnly) {
                # ── 成功 ──
                $resultNote = $item.ExtNote
                if ($isWarningOnly) {
                    $warnMsg    = ($stderr.Trim() -replace '[\r\n]+', ' ')
                    $resultNote = if ($resultNote) { "$resultNote; Warn: $warnMsg" } else { "Warn: $warnMsg" }
                }
                $successList.Add([pscustomobject]@{
                    JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                    MatchMethod = $item.MatchMethod; DestPath = $item.DestPath
                    ExtFixed = $item.ExtFixed; Args = $item.ArgString
                })
                Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                    -MatchMethod $item.MatchMethod -Result 'Updated' `
                    -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                    -HasGps $item.LogHasGps -Description $item.LogDesc `
                    -DestPath $item.DestPath -Notes $resultNote)

                # ── _original バックアップファイルの処理（OutputPath 未指定時のみ） ──
                if (-not $useOutputPath -and -not $NoBackup) {
                    $origFile = $item.MediaPath + '_original'
                    if (Test-Path -LiteralPath $origFile) {
                        switch ($OriginalFileAction) {
                            'Delete' {
                                Remove-Item -LiteralPath $origFile -Force -ErrorAction SilentlyContinue
                            }
                            'Rename' {
                                $origExt  = [System.IO.Path]::GetExtension($item.MediaPath)
                                $origBase = [System.IO.Path]::GetFileNameWithoutExtension($item.MediaPath)
                                $origDir  = Split-Path -Parent $item.MediaPath
                                $newName  = Join-Path $origDir "${origBase}_original${origExt}"
                                if (-not (Test-Path -LiteralPath $newName)) {
                                    Rename-Item -LiteralPath $origFile -NewName (Split-Path -Leaf $newName) -Force -ErrorAction SilentlyContinue
                                }
                                else {
                                    Remove-Item -LiteralPath $origFile -Force -ErrorAction SilentlyContinue
                                }
                            }
                            'Keep' { }
                        }
                    }
                }
            }
            else {
                # ── ExifTool エラー ──
                $errMsg = ($stderr.Trim() -replace '[\r\n]+', ' ')
                $failExifTool.Add([pscustomobject]@{
                    JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                    ExitCode = $p.ExitCode; Error = $errMsg
                })
                Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                    -MatchMethod $item.MatchMethod -Result 'Error' `
                    -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                    -HasGps $item.LogHasGps -Description $item.LogDesc `
                    -DestPath $item.DestPath -Notes "Exit $($p.ExitCode): $errMsg")

                # ExifTool 失敗時もコピー先ファイルは保持
                if ($useOutputPath -and (Test-Path -LiteralPath $item.DestPath)) {
                    Write-Host "    ExifTool失敗: メタデータなしで保持 $($item.DestPath)" -ForegroundColor Yellow
                }
            }
        }
        catch {
            $failExifTool.Add([pscustomobject]@{
                JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                ExitCode = -1; Error = $_.Exception.Message
            })
            Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                -MatchMethod $item.MatchMethod -Result 'Error' `
                -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                -HasGps $item.LogHasGps -Description $item.LogDesc `
                -DestPath $item.DestPath -Notes "Exception: $($_.Exception.Message)")

            if ($useOutputPath -and $item.DestPath -and (Test-Path -LiteralPath $item.DestPath)) {
                Write-Host "    例外発生: メタデータなしで保持 $($item.DestPath)" -ForegroundColor Yellow
            }
        }
        finally {
            if ($argFilePath -and (Test-Path -LiteralPath $argFilePath)) {
                Remove-Item -LiteralPath $argFilePath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Progress -Activity "ExifTool 実行中" -Status "完了" -PercentComplete 100 -Completed
}
else {
    # ═══════ 並列実行（RunspacePool） ═══════
    Write-Host "`nPhase 3: 並列実行（$totalWork 件, $Threads スレッド）..."

    # ── 実行スクリプトブロック（各 Runspace で独立実行） ──
    $executeBlock = {
        param(
            [hashtable]$Item,
            [string]$ExifToolPath,
            [bool]$NoBackup,
            [bool]$UseOutputPath,
            [string]$OriginalFileAction
        )

        $result = @{
            Index      = $Item.Index
            Status     = ''
            ExitCode   = 0
            ErrorMsg   = ''
            IsWarning  = $false
            WarningMsg = ''
        }

        $targetFile = $Item.MediaPath

        # ── OutputPath へのコピー ──
        if ($UseOutputPath -and $Item.DestPath) {
            try {
                Copy-Item -LiteralPath $Item.MediaPath -Destination $Item.DestPath -Force -ErrorAction Stop
                $targetFile = $Item.DestPath
            }
            catch {
                $result.Status   = 'CopyError'
                $result.ErrorMsg = "Copy failed: $($_.Exception.Message)"
                return $result
            }
        }

        # ── ExifTool 引数ファイル構築 ──
        $cmdArgs = [System.Collections.Generic.List[string]]::new()
        if ($NoBackup -or $UseOutputPath) { $cmdArgs.Add('-overwrite_original') }
        $cmdArgs.Add('-P')
        $cmdArgs.Add('-charset')
        $cmdArgs.Add('filename=utf8')
        foreach ($a in $Item.ExifArgs) { $cmdArgs.Add($a) }
        $cmdArgs.Add($targetFile)

        $argFilePath = [System.IO.Path]::GetTempFileName()
        try {
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllLines($argFilePath, $cmdArgs.ToArray(), $utf8NoBom)

            $si = New-Object System.Diagnostics.ProcessStartInfo
            $si.FileName               = $ExifToolPath
            $si.Arguments              = "-@ `"$argFilePath`""
            $si.RedirectStandardOutput = $true
            $si.RedirectStandardError  = $true
            $si.UseShellExecute        = $false
            $si.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $si.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

            $p      = [System.Diagnostics.Process]::Start($si)
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $p.WaitForExit()

            $isWarn = ($p.ExitCode -ne 0 -and $stderr -match 'Warning:' -and
                       $stderr -notmatch 'Error:' -and $stdout -match '1 image files updated')

            if ($p.ExitCode -eq 0 -or $isWarn) {
                $result.Status = 'Success'
                if ($isWarn) {
                    $result.IsWarning  = $true
                    $result.WarningMsg = ($stderr.Trim() -replace '[\r\n]+', ' ')
                }
            }
            else {
                $result.Status   = 'ExifToolError'
                $result.ExitCode = $p.ExitCode
                $result.ErrorMsg = ($stderr.Trim() -replace '[\r\n]+', ' ')
            }

            # ── _original ファイルの処理（非 OutputPath 時） ──
            if ($result.Status -eq 'Success' -and -not $UseOutputPath -and -not $NoBackup) {
                $origFile = $targetFile + '_original'
                if (Test-Path -LiteralPath $origFile) {
                    switch ($OriginalFileAction) {
                        'Delete' {
                            Remove-Item -LiteralPath $origFile -Force -ErrorAction SilentlyContinue
                        }
                        'Rename' {
                            $oExt  = [System.IO.Path]::GetExtension($targetFile)
                            $oBase = [System.IO.Path]::GetFileNameWithoutExtension($targetFile)
                            $oDir  = Split-Path -Parent $targetFile
                            $newN  = Join-Path $oDir "${oBase}_original${oExt}"
                            if (-not (Test-Path -LiteralPath $newN)) {
                                Rename-Item -LiteralPath $origFile -NewName (Split-Path -Leaf $newN) -Force -ErrorAction SilentlyContinue
                            }
                            else {
                                Remove-Item -LiteralPath $origFile -Force -ErrorAction SilentlyContinue
                            }
                        }
                        'Keep' { }
                    }
                }
            }
        }
        catch {
            $result.Status   = 'Exception'
            $result.ExitCode = -1
            $result.ErrorMsg = $_.Exception.Message
        }
        finally {
            if ($argFilePath -and (Test-Path -LiteralPath $argFilePath)) {
                Remove-Item -LiteralPath $argFilePath -Force -ErrorAction SilentlyContinue
            }
        }

        return $result
    }

    # ── RunspacePool 初期化 ──
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $pool.Open()

    # ── タスク投入 ──
    $runspaces = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($item in $workItems) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($executeBlock)
        [void]$ps.AddArgument($item)
        [void]$ps.AddArgument($ExifToolPath)
        [void]$ps.AddArgument([bool]$NoBackup)
        [void]$ps.AddArgument($useOutputPath)
        [void]$ps.AddArgument($OriginalFileAction)

        $runspaces.Add(@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
            Item       = $item
        })
    }

    # ── 結果回収（完了順にポーリング） ──
    $completedCount = 0
    $remaining = [System.Collections.Generic.List[hashtable]]::new($runspaces)

    while ($remaining.Count -gt 0) {
        for ($ri = $remaining.Count - 1; $ri -ge 0; $ri--) {
            $r = $remaining[$ri]
            if (-not $r.Handle.IsCompleted) { continue }

            # ── 完了したタスクの結果を回収 ──
            $execResult = $r.PowerShell.EndInvoke($r.Handle)
            $r.PowerShell.Dispose()
            $remaining.RemoveAt($ri)
            $completedCount++

            Write-Progress -Activity "ExifTool 並列実行中" `
                -Status "$completedCount / $totalWork" `
                -PercentComplete ([int](100 * $completedCount / $totalWork))

            $item = $r.Item
            $res  = if ($execResult -and $execResult.Count -gt 0) { $execResult[0] }
                    else { @{ Status = 'Exception'; ExitCode = -1; ErrorMsg = 'No result from runspace' } }

            # ── 結果に基づいてログ・リスト更新 ──
            if ($res.Status -eq 'Success') {
                $resultNote = $item.ExtNote
                if ($res.IsWarning) {
                    $resultNote = if ($resultNote) { "$resultNote; Warn: $($res.WarningMsg)" } else { "Warn: $($res.WarningMsg)" }
                }
                $successList.Add([pscustomobject]@{
                    JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                    MatchMethod = $item.MatchMethod; DestPath = $item.DestPath
                    ExtFixed = $item.ExtFixed; Args = $item.ArgString
                })
                Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                    -MatchMethod $item.MatchMethod -Result 'Updated' `
                    -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                    -HasGps $item.LogHasGps -Description $item.LogDesc `
                    -DestPath $item.DestPath -Notes $resultNote)
            }
            else {
                # CopyError / ExifToolError / Exception
                $failExifTool.Add([pscustomobject]@{
                    JsonPath = $item.JsonPath; MediaPath = $item.MediaPath
                    ExitCode = $res.ExitCode; Error = $res.ErrorMsg
                })
                $errDetail = switch ($res.Status) {
                    'CopyError' { $res.ErrorMsg }
                    'Exception' { "Exception: $($res.ErrorMsg)" }
                    default     { "Exit $($res.ExitCode): $($res.ErrorMsg)" }
                }
                Write-LogLine (New-LogFields -JsonPath $item.JsonPath -MediaPath $item.MediaPath `
                    -MatchMethod $item.MatchMethod -Result 'Error' `
                    -PhotoTime $item.LogPhotoTime -CreateTime $item.LogCreateTime `
                    -HasGps $item.LogHasGps -Description $item.LogDesc `
                    -DestPath $item.DestPath -Notes $errDetail)

                if ($useOutputPath -and $item.DestPath -and (Test-Path -LiteralPath $item.DestPath)) {
                    Write-Host "    ExifTool失敗: メタデータなしで保持 $($item.DestPath)" -ForegroundColor Yellow
                }
            }
        }

        if ($remaining.Count -gt 0) {
            Start-Sleep -Milliseconds 50
        }
    }

    $pool.Close()
    $pool.Dispose()
    Write-Progress -Activity "ExifTool 並列実行中" -Status "完了" -PercentComplete 100 -Completed
}

# ══════════════════════════════════════════════════════════
# Phase 4: レポート出力
# ══════════════════════════════════════════════════════════

$elapsed        = (Get-Date) - $startTime
$totalProcessed = $jsonFiles.Count
$totalUpdated   = $successList.Count
$totalErrors    = $failExifTool.Count
$totalNoMedia   = $failNoMedia.Count
$totalNoMeta    = $failNoMeta.Count

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host " 処理結果サマリー" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "処理対象 JSON 数:  $totalProcessed"
Write-Host "更新成功/予定:     $totalUpdated" -ForegroundColor Green
$extFixedCount = @($successList | Where-Object { $_.ExtFixed }).Count
if ($extFixedCount -gt 0) {
    Write-Host "  拡張子修正:      $extFixedCount 件" -ForegroundColor Cyan
}
Write-Host "メディア不一致:    $totalNoMedia" -ForegroundColor Yellow
Write-Host "メタデータなし:    $totalNoMeta" -ForegroundColor Yellow
Write-Host "JSON解析エラー:    $($failJsonParse.Count)" -ForegroundColor Red
Write-Host "ExifToolエラー:    $totalErrors" -ForegroundColor Red
if ($Threads -gt 1) {
    Write-Host "実行スレッド数:    $Threads" -ForegroundColor Cyan
}
if ($useOutputPath) {
    Write-Host "出力先フォルダ:    $OutputPath" -ForegroundColor Cyan
}
Write-Host "処理時間:          $($elapsed.ToString('hh\:mm\:ss\.ff'))" -ForegroundColor Cyan
Write-Host "ログファイル:      $LogFile"
Write-Host ""

# ── 失敗詳細レポート ──
$reportPath  = [System.IO.Path]::ChangeExtension($LogFile, '.failure-report.txt')
$reportLines = [System.Collections.Generic.List[string]]::new()

$reportLines.Add("=== 失敗・結果レポート ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) ===")
$reportLines.Add("対象パス: $PhotosPath")
if ($useOutputPath) { $reportLines.Add("出力先:   $OutputPath (年/月階層)") }
if ($Threads -gt 1) { $reportLines.Add("スレッド: $Threads") }
$reportLines.Add("モード:   $(if($WhatIf){'WhatIf (ドライラン)'}else{'実行'})")
$reportLines.Add("処理時間: $($elapsed.ToString('hh\:mm\:ss\.ff'))")
$reportLines.Add("")

# [1] JSON 解析エラー
$reportLines.Add("─── [1] JSON 解析エラー ($($failJsonParse.Count) 件) ───")
if ($failJsonParse.Count -gt 0) {
    foreach ($fp in $failJsonParse) { $reportLines.Add("  $fp") }
}
else { $reportLines.Add("  (なし)") }
$reportLines.Add("")

# [2] メディアファイル不一致
$reportLines.Add("─── [2] メディアファイル不一致 ($($failNoMedia.Count) 件) ───")
if ($failNoMedia.Count -gt 0) {
    foreach ($item in $failNoMedia) {
        $reportLines.Add("  JSON:   $($item.JsonPath)")
        $reportLines.Add("  Title:  $($item.Title)")
        $reportLines.Add("  Method: $($item.MatchAttempted)")
        $reportLines.Add("")
    }
}
else { $reportLines.Add("  (なし)") }
$reportLines.Add("")

# [3] メタデータなし
$reportLines.Add("─── [3] 書き込むメタデータなし ($($failNoMeta.Count) 件) ───")
if ($failNoMeta.Count -gt 0) {
    foreach ($fp in $failNoMeta) { $reportLines.Add("  $fp") }
}
else { $reportLines.Add("  (なし)") }
$reportLines.Add("")

# [4] ExifTool エラー
$reportLines.Add("─── [4] ExifTool エラー ($($failExifTool.Count) 件) ───")
if ($failExifTool.Count -gt 0) {
    foreach ($item in $failExifTool) {
        $reportLines.Add("  JSON:   $($item.JsonPath)")
        $reportLines.Add("  Media:  $($item.MediaPath)")
        $reportLines.Add("  Exit:   $($item.ExitCode)")
        $reportLines.Add("  Error:  $($item.Error)")
        $reportLines.Add("")
    }
}
else { $reportLines.Add("  (なし)") }
$reportLines.Add("")

# [5] 成功 / WhatIf
if ($WhatIf) {
    $reportLines.Add("─── [5] WhatIf: 変換予定 ($($successList.Count) 件) ───")
    $methodStats = $successList | Group-Object -Property MatchMethod | Sort-Object Count -Descending
    $reportLines.Add("  マッチ方法別統計:")
    foreach ($g in $methodStats) {
        $reportLines.Add("    $($g.Name): $($g.Count) 件")
    }
    $reportLines.Add("")
    $reportLines.Add("  ファイルリスト:")
    foreach ($s in $successList) {
        if ($useOutputPath -and $s.DestPath) {
            $reportLines.Add("    [$($s.MatchMethod)] $($s.MediaPath) → $($s.DestPath)")
        }
        else {
            $reportLines.Add("    [$($s.MatchMethod)] $($s.MediaPath)")
        }
    }
}
else {
    $reportLines.Add("─── [5] 更新成功 ($($successList.Count) 件) ───")
    foreach ($s in $successList) {
        if ($useOutputPath -and $s.DestPath) {
            $reportLines.Add("    [$($s.MatchMethod)] $($s.MediaPath) → $($s.DestPath)")
        }
        else {
            $reportLines.Add("    [$($s.MatchMethod)] $($s.MediaPath)")
        }
    }
}

$reportContent = $reportLines -join "`r`n"
Set-Content -Path $reportPath -Value $reportContent -Encoding UTF8

Write-Host "失敗レポート: $reportPath"

# ── WhatIf 時の追加コンソール出力 ──
if ($WhatIf) {
    Write-Host ""
    Write-Host "=== WhatIf マッチ方法統計 ===" -ForegroundColor Magenta
    $successList | Group-Object -Property MatchMethod | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,-30} {1,6} 件" -f $_.Name, $_.Count)
    }
    Write-Host ""

    # 出力先フォルダ使用時: 年/月 分布統計
    if ($useOutputPath -and $successList.Count -gt 0) {
        Write-Host "=== WhatIf 出力先フォルダ分布 ===" -ForegroundColor Cyan
        $destStats = $successList | Where-Object { $_.DestPath } | ForEach-Object {
            $rel = $_.DestPath.Substring($OutputPath.Length).TrimStart('\', '/')
            $parts = $rel -split '[/\\]'
            if ($parts.Count -ge 2) { "$($parts[0])/$($parts[1])" } else { $parts[0] }
        } | Group-Object | Sort-Object Name
        foreach ($g in $destStats) {
            Write-Host ("  {0,-15} {1,6} 件" -f $g.Name, $g.Count)
        }
        Write-Host ""
    }

    $totalFailures = $failJsonParse.Count + $failNoMedia.Count + $failNoMeta.Count
    if ($totalFailures -gt 0) {
        Write-Host "=== WhatIf 失敗サマリー ($totalFailures 件) ===" -ForegroundColor Red
        if ($failJsonParse.Count -gt 0) {
            Write-Host "  JSON解析エラー: $($failJsonParse.Count) 件" -ForegroundColor Red
        }
        if ($failNoMedia.Count -gt 0) {
            Write-Host "  メディア不一致: $($failNoMedia.Count) 件" -ForegroundColor Yellow
            Write-Host "    (上位5件):" -ForegroundColor Yellow
            $failNoMedia | Select-Object -First 5 | ForEach-Object {
                Write-Host "      Title=$($_.Title)  JSON=$($_.JsonPath)" -ForegroundColor Yellow
            }
        }
        if ($failNoMeta.Count -gt 0) {
            Write-Host "  メタデータなし: $($failNoMeta.Count) 件" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "詳細は失敗レポートを参照: $reportPath" -ForegroundColor Yellow
    }
    else {
        Write-Host "すべてのデータが正常に変換される見込みです。" -ForegroundColor Green
    }
}
