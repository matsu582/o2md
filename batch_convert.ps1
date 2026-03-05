<#
.SYNOPSIS
    指定フォルダ配下のMS Officeファイルをo2mdで一括Markdown変換するスクリプト

.DESCRIPTION
    入力フォルダ内のOfficeファイル（Excel, Word, PowerPoint, PDF）を再帰的に検索し、
    フォルダ構成を維持したまま出力フォルダへMarkdownファイルとして変換します。
    変換にはo2md.pyを使用します。

.PARAMETER InputDir
    変換対象のOfficeファイルが格納されたフォルダパス（必須）

.PARAMETER OutputDir
    変換結果の出力先フォルダパス（省略時: ./batch_output）

.PARAMETER O2mdDir
    o2mdリポジトリのディレクトリパス（省略時: スクリプトと同じフォルダ）

.PARAMETER Format
    出力画像形式。svg または png を指定（省略時: svg）

.PARAMETER UseHeadingText
    [Word専用] 見出しテキストをリンクに使用する場合に指定

.PARAMETER ShapeMetadata
    図形メタデータを出力する場合に指定

.PARAMETER OcrEngine
    [PDF専用] OCRエンジンを指定（tesseract または manga-ocr、省略時: tesseract）

.PARAMETER TessdataDir
    [PDF専用] tessdataディレクトリのパスを指定

.PARAMETER Docling
    [PDF専用] doclingによる表検出を有効にする場合に指定

.PARAMETER Verbose
    詳細なデバッグ出力を表示する場合に指定

.EXAMPLE
    # 基本的な使用方法
    .\batch_convert.ps1 -InputDir "C:\Documents\Office Files"

.EXAMPLE
    # 出力先とo2mdディレクトリを指定
    .\batch_convert.ps1 -InputDir "C:\docs" -OutputDir "C:\converted" -O2mdDir "C:\tools\o2md"

.EXAMPLE
    # PNG形式で出力、Word文書は見出しテキストリンクを使用
    .\batch_convert.ps1 -InputDir "C:\docs" -Format png -UseHeadingText
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "変換対象のフォルダパスを指定してください")]
    [string]$InputDir,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = ".\batch_output",

    [Parameter(Mandatory = $false)]
    [string]$O2mdDir = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("svg", "png")]
    [string]$Format = "svg",

    [Parameter(Mandatory = $false)]
    [switch]$UseHeadingText,

    [Parameter(Mandatory = $false)]
    [switch]$ShapeMetadata,

    [Parameter(Mandatory = $false)]
    [ValidateSet("tesseract", "manga-ocr")]
    [string]$OcrEngine = "tesseract",

    [Parameter(Mandatory = $false)]
    [string]$TessdataDir = "",

    [Parameter(Mandatory = $false)]
    [switch]$Docling,

    [Parameter(Mandatory = $false)]
    [switch]$Verbose
)

# --- 定数定義 ---
# 対応するOfficeファイルの拡張子一覧
$SUPPORTED_EXTENSIONS = @("*.xlsx", "*.xls", "*.docx", "*.doc", "*.pptx", "*.ppt", "*.pdf")

# --- 関数定義 ---

function Write-Log {
    <#
    .SYNOPSIS
        タイムスタンプ付きのログメッセージを出力する
    #>
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-UvAvailable {
    <#
    .SYNOPSIS
        uvコマンドが利用可能か確認する
    #>
    try {
        $null = Get-Command "uv" -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Get-OfficeFiles {
    <#
    .SYNOPSIS
        指定フォルダから対応するOfficeファイルを再帰的に取得する
    #>
    param(
        [string]$SearchDir,
        [string[]]$Extensions
    )
    $allFiles = @()
    foreach ($ext in $Extensions) {
        $found = Get-ChildItem -Path $SearchDir -Filter $ext -Recurse -File -ErrorAction SilentlyContinue
        if ($found) {
            $allFiles += $found
        }
    }
    return $allFiles | Sort-Object FullName
}

function Build-O2mdArguments {
    <#
    .SYNOPSIS
        o2md.pyに渡すコマンドライン引数を組み立てる
    #>
    param(
        [string]$FilePath,
        [string]$FileOutputDir
    )

    $argList = @($FilePath, "-o", $FileOutputDir, "--format", $Format)

    if ($UseHeadingText) {
        $argList += "--use-heading-text"
    }
    if ($ShapeMetadata) {
        $argList += "--shape-metadata"
    }
    if ($OcrEngine -ne "tesseract") {
        $argList += @("--ocr-engine", $OcrEngine)
    }
    if ($TessdataDir -ne "") {
        $argList += @("--tessdata-dir", $TessdataDir)
    }
    if ($Docling) {
        $argList += "--docling"
    }
    if ($Verbose) {
        $argList += "-v"
    }

    return $argList
}

function Invoke-O2mdConversion {
    <#
    .SYNOPSIS
        1つのファイルに対してo2md変換を実行する
    .OUTPUTS
        変換が成功した場合は$true、失敗した場合は$false
    #>
    param(
        [string]$FilePath,
        [string]$FileOutputDir,
        [string]$ScriptPath
    )

    $o2mdArgs = Build-O2mdArguments -FilePath $FilePath -FileOutputDir $FileOutputDir
    $uvArgs = @("run", "python", $ScriptPath) + $o2mdArgs

    # 一時ファイルのパスを事前に取得
    $stdoutTempFile = [System.IO.Path]::GetTempFileName()
    $stderrTempFile = [System.IO.Path]::GetTempFileName()

    try {
        $proc = Start-Process -FilePath "uv" `
            -ArgumentList $uvArgs `
            -WorkingDirectory $resolvedO2mdDir `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdoutTempFile `
            -RedirectStandardError $stderrTempFile

        # 標準出力の内容を表示
        if (Test-Path $stdoutTempFile) {
            $stdoutContent = Get-Content $stdoutTempFile -Raw
            if ($stdoutContent) {
                Write-Host $stdoutContent
            }
        }

        if ($proc.ExitCode -ne 0) {
            # 標準エラー出力の内容を表示
            if (Test-Path $stderrTempFile) {
                $stderrContent = Get-Content $stderrTempFile -Raw
                if ($stderrContent) {
                    Write-Log "エラー出力: $stderrContent" "ERROR"
                }
            }
            return $false
        }
        return $true
    }
    catch {
        Write-Log "実行例外: $_" "ERROR"
        return $false
    }
    finally {
        # 一時ファイルのクリーンアップ
        Remove-Item $stdoutTempFile -ErrorAction SilentlyContinue
        Remove-Item $stderrTempFile -ErrorAction SilentlyContinue
    }
}

# --- メイン処理 ---

# 入力フォルダの存在確認
$resolvedInputDir = Resolve-Path -Path $InputDir -ErrorAction SilentlyContinue
if (-not $resolvedInputDir) {
    Write-Log "入力フォルダが見つかりません: $InputDir" "ERROR"
    exit 1
}
$resolvedInputDir = $resolvedInputDir.Path

# o2mdディレクトリの解決
if ($O2mdDir -eq "") {
    $resolvedO2mdDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $resolvedO2mdDir = Resolve-Path -Path $O2mdDir -ErrorAction SilentlyContinue
    if (-not $resolvedO2mdDir) {
        Write-Log "o2mdディレクトリが見つかりません: $O2mdDir" "ERROR"
        exit 1
    }
    $resolvedO2mdDir = $resolvedO2mdDir.Path
}

# o2md.pyの存在確認
$o2mdScript = Join-Path $resolvedO2mdDir "o2md.py"
if (-not (Test-Path $o2mdScript)) {
    Write-Log "o2md.pyが見つかりません: $o2mdScript" "ERROR"
    Write-Log "O2mdDirパラメータでo2mdリポジトリのパスを指定してください" "ERROR"
    exit 1
}

# uvコマンドの確認
if (-not (Test-UvAvailable)) {
    Write-Log "uvコマンドが見つかりません。uvをインストールしてください。" "ERROR"
    Write-Log "インストール方法: powershell -ExecutionPolicy ByPass -c `"irm https://astral.sh/uv/install.ps1 | iex`"" "ERROR"
    exit 1
}

# 出力フォルダの作成
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not (Test-Path $resolvedOutputDir)) {
    New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
    Write-Log "出力フォルダを作成しました: $resolvedOutputDir"
}

# 対象ファイルの検索
Write-Log "対象ファイルを検索中: $resolvedInputDir"
$targetFiles = Get-OfficeFiles -SearchDir $resolvedInputDir -Extensions $SUPPORTED_EXTENSIONS

if ($targetFiles.Count -eq 0) {
    Write-Log "変換対象のOfficeファイルが見つかりませんでした" "WARNING"
    exit 0
}

Write-Log "変換対象ファイル数: $($targetFiles.Count) 件"
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " o2md 一括変換を開始します" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  入力フォルダ : $resolvedInputDir"
Write-Host "  出力フォルダ : $resolvedOutputDir"
Write-Host "  o2mdディレクトリ: $resolvedO2mdDir"
Write-Host "  画像形式     : $Format"
Write-Host "  対象ファイル数: $($targetFiles.Count) 件"
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 変換処理の実行
$successCount = 0
$failCount = 0
$failedFiles = @()

foreach ($file in $targetFiles) {
    # 入力フォルダからの相対パスを算出
    $relativePath = $file.FullName.Substring($resolvedInputDir.Length).TrimStart('\', '/')
    $relativeDir = Split-Path -Parent $relativePath

    # 出力先フォルダ（入力フォルダ構成を維持）
    $fileOutputDir = $resolvedOutputDir
    if ($relativeDir -ne "") {
        $fileOutputDir = Join-Path $resolvedOutputDir $relativeDir
    }

    # 出力先サブフォルダが無ければ作成
    if (-not (Test-Path $fileOutputDir)) {
        New-Item -ItemType Directory -Path $fileOutputDir -Force | Out-Null
    }

    $currentIndex = $successCount + $failCount + 1
    Write-Host ""
    Write-Log "[$currentIndex/$($targetFiles.Count)] 変換中: $relativePath"

    # o2md変換の実行
    $result = Invoke-O2mdConversion `
        -FilePath $file.FullName `
        -FileOutputDir $fileOutputDir `
        -ScriptPath $o2mdScript

    if ($result) {
        $successCount++
        Write-Log "変換成功: $relativePath" "SUCCESS"
    }
    else {
        $failCount++
        $failedFiles += $relativePath
        Write-Log "変換失敗: $relativePath" "ERROR"
    }
}

# --- 結果サマリーの表示 ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " 一括変換が完了しました" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  成功: $successCount 件" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "  失敗: $failCount 件" -ForegroundColor Red
    Write-Host ""
    Write-Host "  失敗したファイル:" -ForegroundColor Red
    foreach ($failedFile in $failedFiles) {
        Write-Host "    - $failedFile" -ForegroundColor Red
    }
}
else {
    Write-Host "  失敗: 0 件"
}
Write-Host "  出力先: $resolvedOutputDir"
Write-Host "========================================" -ForegroundColor Cyan

# 失敗があった場合は終了コード1で終了
if ($failCount -gt 0) {
    exit 1
}
exit 0
