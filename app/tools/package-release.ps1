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
    $python = Get-Command py -ErrorAction SilentlyContinue
    if ($python) { & $python.Source -3 $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { & $python.Source $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
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

function Remove-SharedFolderDevelopmentFiles([string]$StageRoot) {
    foreach ($relativePath in @(
        '.git',
        '.github',
        '.agents',
        '.codex',
        '.gitignore',
        'CHANGELOG_V4.md',
        'CHANGELOG_V5.md',
        '日本語管理.vbs',
        '英語管理.vbs',
        '共有フォルダー用フォルダー作成.cmd',
        'docs\API.md',
        'docs\ReportBinder_UIUX改修指示書_V4.md',
        'app\lib\pdfbox\src',
        'app\lib\pdfbox\build.ps1',
        'app\tools\fixtures',
        'app\tools\selfcheck.py',
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
        '日本語管理.cmd',
        '英語管理.cmd',
        'README.md',
        'THIRD_PARTY_NOTICES.md',
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
        'app\tools\package-release.ps1'
    )
    $remaining = @($forbidden | Where-Object {
        Test-Path -LiteralPath (Join-Path $StageRoot $_)
    })
    if ($remaining.Count -gt 0) {
        throw ("共有フォルダー用配布物に不要なファイルが残っています:`n" + ($remaining -join "`n"))
    }
}

function Copy-ReleaseSource([string]$StageRoot) {
    foreach ($item in (Get-ChildItem -LiteralPath $root -Force)) {
        if ($item.Name -eq '.git') { continue }
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
        Remove-SharedFolderDevelopmentFiles $stage
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
        Copy-Item -Path (Join-Path $root '*') -Destination $stage -Recurse -Force
        Remove-RuntimeFiles $stage
        if (-not $IncludeJava) {
            $java = Join-Path $stage 'app\lib\java'
            if (Test-Path -LiteralPath $java) { Remove-Item -LiteralPath $java -Recurse -Force }
        }
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
