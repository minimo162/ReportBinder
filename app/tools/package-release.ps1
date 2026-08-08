param(
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder\release'),
    [Parameter(Mandatory=$false)]
    [switch]$SharedFolderOnly,
    [Parameter(Mandatory=$false)]
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$rootPrefix = $root
if (-not $rootPrefix.EndsWith([IO.Path]::DirectorySeparatorChar)) {
    $rootPrefix += [IO.Path]::DirectorySeparatorChar
}
if ($OutputDir -eq $root -or $OutputDir.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw '出力先にはReportBinderの元フォルダー配下を指定できません。'
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

function Invoke-SelfCheck {
    $script = Join-Path $root 'app\tools\selfcheck.py'
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { & $python.Source $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
    $python = Get-Command py -ErrorAction SilentlyContinue
    if ($python) { & $python.Source -3 $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
    throw 'Pythonが見つかりません。先に app\tools\selfcheck.py を実行してください。'
}

function Invoke-ThirdPartyCheck([bool]$RequirePortableJava) {
    $script = Join-Path $root 'app\tools\verify-thirdparty.ps1'
    if (-not (Test-Path -LiteralPath $script)) {
        throw "依存物検証スクリプトが見つかりません: $script"
    }
    if ($RequirePortableJava) {
        & $script -RequirePdfJs -RequirePortableJava
    } else {
        & $script -RequirePdfJs
    }
    if ($LASTEXITCODE -ne 0) {
        throw '第三者依存物の検証に失敗しました。app\tools\install-thirdparty.cmd を実行してください。'
    }
}

function Assert-StagedDependencies([string]$StageRoot, [bool]$IncludeJava) {
    $required = @(
        (Join-Path $StageRoot 'app\lib\pdfbox\pdfbox-app.jar'),
        (Join-Path $StageRoot 'app\lib\pdfbox\ReportPdfComposer.jar'),
        (Join-Path $StageRoot 'app\web\pdfjs\pdf.min.mjs'),
        (Join-Path $StageRoot 'app\web\pdfjs\pdf.worker.min.mjs'),
        (Join-Path $StageRoot 'app\web\pdfjs\LICENSE')
    )
    if ($IncludeJava) {
        $required += (Join-Path $StageRoot 'app\lib\java\bin\java.exe')
        $required += (Join-Path $StageRoot 'app\lib\java\release')
        $required += (Join-Path $StageRoot 'app\lib\java\NOTICE')
        $required += (Join-Path $StageRoot 'app\lib\java\JAVA_VERSION.txt')
    }
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        throw ("配布物に必要な依存ファイルがありません:`n" + ($missing -join "`n"))
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $composer = [IO.Compression.ZipFile]::OpenRead((Join-Path $StageRoot 'app\lib\pdfbox\ReportPdfComposer.jar'))
    try {
        foreach ($entry in @('ReportPdfComposer.class','ReportPdfComposer$ContentSpec.class','BatchPdfSplitter.class','PdfPageAnalyzer.class','PdfBatchRasterizer.class')) {
            if ($null -eq $composer.GetEntry($entry)) { throw "配布物のPDF組版jarに必要なクラスがありません: $entry" }
        }
    } finally { $composer.Dispose() }
}

function Remove-PathIfExists([string]$Path) {
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Remove-RuntimeFiles([string]$StageRoot) {
    foreach ($path in @(
        (Join-Path $StageRoot 'app\config.json'),
        (Join-Path $StageRoot 'app\thirdparty-cache'),
        (Join-Path $StageRoot 'app\logs'),
        (Join-Path $StageRoot 'app\lib\pdfbox\classes')
    )) {
        Remove-PathIfExists $path
    }
    New-Item -ItemType Directory -Path (Join-Path $StageRoot 'app\logs') -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $StageRoot 'app\logs\.gitkeep') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $StageRoot 'app\thirdparty-cache') -Force | Out-Null
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -File | Where-Object {
        $_.Name -like 'job_*.json' -or $_.Name -like '~building_*.pdf' -or
        $_.Name -like '*.tmp' -or $_.Name -like '*.pyc'
    } | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -Directory |
        Where-Object { $_.Name -eq '__pycache__' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -Directory |
        Where-Object { $_.Name -eq '_reportbinder' -or $_.Name -eq '出力' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function Remove-ReleaseDevelopmentFiles([string]$StageRoot) {
    foreach ($relativePath in @(
        '.git', '.github', '.agents', '.codex', '.gitignore',
        'app\lib\pdfbox\src', 'app\lib\pdfbox\build.ps1',
        'app\tools\fixtures', 'app\tools\selfcheck.py',
        'app\tools\schema-v3-selfcheck.ps1', 'app\tools\source-adapter-selfcheck.ps1',
        'app\tools\pdf-source-adapter-selfcheck.ps1', 'app\tools\word-source-adapter-selfcheck.ps1',
        'app\tools\history-generalization-selfcheck.ps1', 'app\tools\final-composition-selfcheck.py',
        'app\tools\create-word-adapter-fixtures.py', 'app\tools\operational-readiness-selfcheck.ps1',
        'app\tools\create-scale-benchmark-fixtures.py', 'app\tools\create-scale-benchmark-workbooks.mjs',
        'app\tools\scale-benchmark.ps1', 'docs\SCALE_BENCHMARK.md', 'docs\benchmarks',
        'app\tools\ci-selfcheck.ps1', 'app\tools\history-logic-selfcheck.ps1',
        'app\tools\create-pdf-diff-corpus.py', 'app\tools\pdf-diff-corpus-selfcheck.ps1',
        'tests\pdf-diff-corpus.mjs', 'docs\PDF_DIFF_CORPUS.md',
        'app\tools\ci-requirements.txt',
        'app\tools\package-release.ps1'
    )) { Remove-PathIfExists (Join-Path $StageRoot $relativePath) }
}

# 完全性マニフェストの除外規則。launch.ps1 に同じ関数があり、両者が一致していることを
# selfcheck.py が検査する。片方だけを変えると検証が素通りするため、必ず同時に直すこと。
function Test-IntegrityExcludedPath([string]$RelativePath) {
    if ($RelativePath -eq 'integrity-manifest.json') { return $true }
    if ($RelativePath -eq 'config.json') { return $true }
    $top = ($RelativePath -split '/')[0]
    return ($top -in @('logs', 'thirdparty-cache'))
}

# 共有フォルダーの配布ツリーは、各利用者のPCへそのままコピーされてから実行される。
# コピー前に照合できるよう、app配下の全ファイルのSHA-256とサイズを記録する。
# launch.ps1 側の検証と対になっているため、除外規則を両者で一致させること。
function Write-IntegrityManifest([string]$StageRoot) {
    $appRoot = Join-Path $StageRoot 'app'
    $prefix = [IO.Path]::GetFullPath($appRoot).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $manifestPath = Join-Path $appRoot 'integrity-manifest.json'
    Remove-PathIfExists $manifestPath
    $runtimeVersion = 'legacy'
    try { $runtimeVersion = [string]((Get-Content -LiteralPath (Join-Path $appRoot 'runtime-version.json') -Raw -Encoding UTF8 | ConvertFrom-Json).version) } catch { }
    $files = [ordered]@{}
    $count = 0
    foreach ($file in @(Get-ChildItem -LiteralPath $appRoot -Recurse -File -Force | Sort-Object FullName)) {
        $relative = $file.FullName.Substring($prefix.Length) -replace '\\', '/'
        if (Test-IntegrityExcludedPath $relative) { continue }
        $files[$relative] = [ordered]@{
            sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            size = [int64]$file.Length
        }
        $count++
    }
    if ($count -lt 1) { throw '完全性マニフェストの対象ファイルが見つかりません。' }
    $manifest = [ordered]@{
        schemaVersion = 1
        product = 'ReportBinder'
        runtimeVersion = $runtimeVersion
        createdAt = (Get-Date).ToString('o')
        fileCount = $count
        files = $files
    }
    [IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    Write-Output ("integrity manifest: {0} files" -f $count)
}

function Write-ReleaseManifest([string]$StageRoot, [string]$Flavor, [bool]$IncludeJava) {
    $runtimeVersion = 'legacy'
    $runtimePath = Join-Path $StageRoot 'app\runtime-version.json'
    if (Test-Path -LiteralPath $runtimePath) {
        try { $runtimeVersion = [string]((Get-Content -LiteralPath $runtimePath -Raw -Encoding UTF8 | ConvertFrom-Json).version) } catch { }
    }
    $composerPath = Join-Path $StageRoot 'app\lib\pdfbox\ReportPdfComposer.jar'
    $pdfboxPath = Join-Path $StageRoot 'app\lib\pdfbox\pdfbox-app.jar'
    $manifest = [ordered]@{
        schemaVersion=1; product='ReportBinder'; flavor=$Flavor; runtimeVersion=$runtimeVersion
        createdAt=(Get-Date).ToString('o'); portableJavaIncluded=$IncludeJava
        sha256=[ordered]@{
            composer=(Get-FileHash -LiteralPath $composerPath -Algorithm SHA256).Hash.ToLowerInvariant()
            pdfbox=(Get-FileHash -LiteralPath $pdfboxPath -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    [IO.File]::WriteAllText((Join-Path $StageRoot 'release-manifest.json'), ($manifest | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
}

function Remove-SharedFolderDevelopmentFiles([string]$StageRoot) {
    foreach ($relativePath in @(
        '.git',
        '.github',
        '.agents',
        '.codex',
        '.gitignore',
        'CHANGELOG_V4.md',
        'CHANGELOG_V5.md',
        '共有フォルダー用フォルダー作成.cmd',
        'docs\API.md',
        'docs\ReportBinder_UIUX改修指示書_V4.md',
        'app\lib\pdfbox\src',
        'app\lib\pdfbox\build.ps1',
        'app\tools\fixtures',
        'app\tools\selfcheck.py',
        'app\tools\create-scale-benchmark-fixtures.py',
        'app\tools\create-scale-benchmark-workbooks.mjs',
        'app\tools\scale-benchmark.ps1',
        'app\tools\ci-selfcheck.ps1',
        'app\tools\history-logic-selfcheck.ps1',
        'app\tools\create-pdf-diff-corpus.py',
        'app\tools\pdf-diff-corpus-selfcheck.ps1',
        'tests\pdf-diff-corpus.mjs',
        'docs\PDF_DIFF_CORPUS.md',
        'app\tools\ci-requirements.txt',
        'docs\SCALE_BENCHMARK.md',
        'docs\benchmarks',
        'app\tools\package-release.ps1'
    )) {
        Remove-PathIfExists (Join-Path $StageRoot $relativePath)
    }
    # 配布先は完成済みのオフライン版なので、空の実行時フォルダーも持ち込まない。
    Remove-PathIfExists (Join-Path $StageRoot 'app\logs')
    Remove-PathIfExists (Join-Path $StageRoot 'app\thirdparty-cache')
}

function Assert-SharedFolderLayout([string]$StageRoot) {
    $required = @(
        '資料をPDFにまとめる.cmd',
        'README.md',
        'THIRD_PARTY_NOTICES.md',
        'release-manifest.json',
        'app\launch.ps1',
        'app\server.ps1',
        'app\default-config.json',
        'app\runtime-version.json',
        'app\web\index.html'
    )
    $missing = @($required | Where-Object {
        -not (Test-Path -LiteralPath (Join-Path $StageRoot $_) -PathType Leaf)
    })
    if ($missing.Count -gt 0) {
        throw ("共有フォルダー用配布物の必須ファイルがありません:`n" + ($missing -join "`n"))
    }

    $forbidden = @(
        '日本語管理.cmd',
        '英語管理.cmd',
        '日本語管理.vbs',
        '英語管理.vbs',
        '共有フォルダー用フォルダー作成.cmd',
        '.git',
        '.github',
        'app\config.json',
        'app\logs',
        'app\thirdparty-cache',
        'app\lib\pdfbox\src',
        'app\lib\pdfbox\build.ps1',
        'app\tools\fixtures',
        'app\tools\selfcheck.py',
        'app\tools\package-release.ps1',
        'app\tools\schema-v3-selfcheck.ps1',
        'app\tools\source-adapter-selfcheck.ps1',
        'app\tools\pdf-source-adapter-selfcheck.ps1',
        'app\tools\word-source-adapter-selfcheck.ps1',
        'app\tools\history-generalization-selfcheck.ps1',
        'app\tools\final-composition-selfcheck.py',
        'app\tools\create-word-adapter-fixtures.py',
        'app\tools\operational-readiness-selfcheck.ps1',
        'app\tools\create-scale-benchmark-fixtures.py',
        'app\tools\create-scale-benchmark-workbooks.mjs',
        'app\tools\scale-benchmark.ps1',
        'app\tools\ci-selfcheck.ps1',
        'app\tools\history-logic-selfcheck.ps1',
        'app\tools\create-pdf-diff-corpus.py',
        'app\tools\pdf-diff-corpus-selfcheck.ps1',
        'tests\pdf-diff-corpus.mjs',
        'docs\PDF_DIFF_CORPUS.md',
        'app\tools\ci-requirements.txt',
        'docs\SCALE_BENCHMARK.md',
        'docs\benchmarks'
    )
    $remaining = @($forbidden | Where-Object {
        Test-Path -LiteralPath (Join-Path $StageRoot $_)
    })
    if ($remaining.Count -gt 0) {
        throw ("共有フォルダー用配布物に不要なファイルが残っています:`n" + ($remaining -join "`n"))
    }
}

function Copy-ReleaseSource([string]$StageRoot) {
    $excludedTopLevelNames = @('.git', '.github', '.agents', '.codex', '.gitignore', 'tmp')
    foreach ($item in (Get-ChildItem -LiteralPath $root -Force)) {
        if ($excludedTopLevelNames -contains $item.Name) { continue }
        Copy-Item -LiteralPath $item.FullName -Destination $StageRoot -Recurse -Force
    }
}

function Get-UniqueReleasePath([string]$BasePath) {
    $candidate = $BasePath
    $index = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = $BasePath + '_' + $index
        $index++
    }
    return $candidate
}

function Copy-DirectoryContents([string]$SourceDir, [string]$DestinationDir) {
    New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    foreach ($item in (Get-ChildItem -LiteralPath $SourceDir -Force)) {
        Copy-Item -LiteralPath $item.FullName -Destination $DestinationDir -Recurse -Force
    }
}

function Publish-SharedFolderStage([string]$StageRoot, [string]$TargetRoot) {
    $moveMessage = ''
    $publishedByMove = $false
    try {
        Move-Item -LiteralPath $StageRoot -Destination $TargetRoot
        $publishedByMove = $true
    } catch {
        $moveMessage = $_.Exception.Message
        # OneDriveやウイルス対策がディレクトリ移動だけを拒否する場合は、
        # 完成済みstageの内容をコピーし、コピー先をもう一度検証する。
        if (-not (Test-Path -LiteralPath $StageRoot) -and (Test-Path -LiteralPath $TargetRoot)) {
            $publishedByMove = $true
        } else {
            if (Test-Path -LiteralPath $TargetRoot) {
                Remove-Item -LiteralPath $TargetRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
            $buildingMarker = Join-Path $TargetRoot '_作成中.txt'
            try {
                New-Item -ItemType Directory -Path $TargetRoot -Force | Out-Null
                New-Item -ItemType File -Path $buildingMarker -Force | Out-Null
                Copy-DirectoryContents -SourceDir $StageRoot -DestinationDir $TargetRoot
            } catch {
                if (Test-Path -LiteralPath $TargetRoot) {
                    Remove-Item -LiteralPath $TargetRoot -Recurse -Force -ErrorAction SilentlyContinue
                }
                throw ("完成フォルダーをコピーできませんでした。フォルダー移動: {0} / 内容コピー: {1}" -f $moveMessage, $_.Exception.Message)
            }
        }
    }

    try {
        Assert-StagedDependencies -StageRoot $TargetRoot -IncludeJava $true
        Assert-SharedFolderLayout $TargetRoot
        $buildingMarker = Join-Path $TargetRoot '_作成中.txt'
        if (Test-Path -LiteralPath $buildingMarker) {
            Remove-Item -LiteralPath $buildingMarker -Force
        }
    } catch {
        if (Test-Path -LiteralPath $TargetRoot) {
            Remove-Item -LiteralPath $TargetRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw
    }

    if (-not $publishedByMove) {
        Write-Warning ("フォルダー移動が拒否されたため、内容コピーで安全に作成しました: {0}" -f $moveMessage)
    }
    return $TargetRoot
}

function New-SharedFolderRelease {
    # OneDrive配下でのstagingは同期処理に掴まれやすいため、OSの一時領域で完成させる。
    $tempBase = Join-Path ([IO.Path]::GetTempPath()) ('ReportBinderShared_' + [guid]::NewGuid().ToString('N'))
    $stage = Join-Path $tempBase 'ReportBinder'
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        Copy-ReleaseSource $stage
        Remove-RuntimeFiles $stage
        Remove-ReleaseDevelopmentFiles $stage
        Remove-SharedFolderDevelopmentFiles $stage
        Write-ReleaseManifest $stage 'shared-folder-offline' $true
        # マニフェストは全ファイル削除が終わった後に作る。順序を変えると、削除済みの
        # ファイルが記録されて利用者側の検証が必ず失敗する。
        Write-IntegrityManifest $stage
        Assert-StagedDependencies -StageRoot $stage -IncludeJava $true
        Assert-SharedFolderLayout $stage

        $targetBase = Join-Path $OutputDir ("ReportBinder_共有フォルダー用_{0}" -f $stamp)
        $target = Get-UniqueReleasePath $targetBase
        return (Publish-SharedFolderStage -StageRoot $stage -TargetRoot $target)
    } finally {
        if (Test-Path -LiteralPath $tempBase) {
            Remove-Item -LiteralPath $tempBase -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-ReleaseZip([string]$Suffix, [bool]$IncludeJava) {
    $tempBase = Join-Path ([IO.Path]::GetTempPath()) "ReportBinderPackage_$([guid]::NewGuid().ToString('N'))"
    $stage = Join-Path $tempBase 'ReportBinder'
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        Copy-ReleaseSource $stage
        Remove-RuntimeFiles $stage
        Remove-ReleaseDevelopmentFiles $stage
        if (-not $IncludeJava) {
            $java = Join-Path $stage 'app\lib\java'
            if (Test-Path -LiteralPath $java) { Remove-Item -LiteralPath $java -Recurse -Force }
        }
        Write-ReleaseManifest $stage $Suffix $IncludeJava
        Write-IntegrityManifest $stage
        Assert-StagedDependencies -StageRoot $stage -IncludeJava $IncludeJava
        $zip = Join-Path $OutputDir "ReportBinder_V5_${Suffix}_${stamp}.zip"
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
        New-Utf8Zip $stage $zip
        return $zip
    } finally {
        if (Test-Path -LiteralPath $tempBase) {
            Remove-Item -LiteralPath $tempBase -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-Utf8Zip([string]$SourceDir, [string]$ZipPath) {
    # Compress-Archive does not reliably mark non-ASCII entry names as UTF-8 for Windows Explorer.
    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $baseParent = Split-Path -Parent $SourceDir
    $fs = [System.IO.File]::Open($ZipPath, [System.IO.FileMode]::Create)
    $archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($file in (Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Force)) {
            $rel = $file.FullName.Substring($baseParent.Length + 1).Replace([char]92, [char]47)
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $file.FullName,
                $rel,
                [System.IO.Compression.CompressionLevel]::Optimal
            )
        }
    } finally {
        $archive.Dispose()
        $fs.Dispose()
    }
}

if ($SharedFolderOnly) {
    # 開発リポジトリ全体を対象にするselfcheckは、ローカルのログやキャッシュでも失敗する。
    # 共有配布作成では、実際に配布するstagingの依存物と完成レイアウトだけを検証する。
    # 共有フォルダー版は、ネットワーク接続なしで各PCへローカル展開できる完成版に限定する。
    Invoke-ThirdPartyCheck $true
    $sharedFolder = New-SharedFolderRelease
    Write-Host "共有フォルダー用の完成フォルダーを作成しました:`n$sharedFolder"
    if (-not $NoOpen) {
        try {
            Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $sharedFolder)
        } catch {
            Write-Warning "エクスプローラーを開けませんでした。作成先を手動で開いてください: $sharedFolder"
        }
    }
    exit 0
}

Invoke-SelfCheck
# PDFBox and PDF.js are included in both ZIP release variants.
Invoke-ThirdPartyCheck $false
# The offline-complete release must contain its own verified portable JRE even if system Java exists.
Invoke-ThirdPartyCheck $true

$offline = New-ReleaseZip 'オフライン完結版' $true
$online  = New-ReleaseZip 'オンライン導入版' $false
Write-Host "作成しました:`n$offline`n$online"
