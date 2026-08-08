param(
    [switch]$SkipPackageSmoke,
    [string]$PackageOutputDir = ''
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$repoRoot = Split-Path -Parent $appRoot
$ownsPackageOutput = [string]::IsNullOrWhiteSpace($PackageOutputDir)
if ($ownsPackageOutput) {
    $PackageOutputDir = Join-Path ([IO.Path]::GetTempPath()) ('ReportBinderCi_' + [Guid]::NewGuid().ToString('N'))
}
$PackageOutputDir = [IO.Path]::GetFullPath($PackageOutputDir)

function Invoke-Checked([string]$Label, [string]$FilePath, [string[]]$Arguments) {
    Write-Output ("== {0} ==" -f $Label)
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Label failed with exit code $LASTEXITCODE." }
}

function Get-PythonCommand {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return [pscustomobject]@{ file=$python.Source; prefix=@() } }
    $python = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($python) { return [pscustomobject]@{ file=$python.Source; prefix=@('-3') } }
    throw 'Python 3 was not found.'
}

function Assert-PowerShellSyntax {
    $errors = New-Object Collections.Generic.List[string]
    # Runtime artifacts and E2E fixtures live under repo/tmp and may be created or
    # removed by a running server while CI is enumerating. Product PowerShell is
    # confined to app/, so keep the syntax walk deterministic and in scope.
    foreach ($file in Get-ChildItem -LiteralPath $appRoot -Recurse -File -Filter '*.ps1') {
        if ($file.FullName -like "$PackageOutputDir*") { continue }
        $tokens = $null
        $parseErrors = $null
        [void][Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
        foreach ($parseError in @($parseErrors)) {
            $relative = $file.FullName.Substring($repoRoot.Length).TrimStart('\')
            $errors.Add(("{0}:{1}: {2}" -f $relative, $parseError.Extent.StartLineNumber, $parseError.Message))
        }
    }
    if ($errors.Count -gt 0) { throw ("PowerShell syntax check failed:`n" + ($errors -join "`n")) }
    Write-Output 'PowerShell syntax check OK.'
}

function Assert-VersionDocumented {
    $runtime = Get-Content -LiteralPath (Join-Path $appRoot 'runtime-version.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $version = ([string]$runtime.version).Trim()
    if ($version -notmatch '^\d{4}\.\d{2}\.\d{2}\.\d+$') { throw "Invalid runtime version: $version" }
    $changelog = Get-Content -LiteralPath (Join-Path $repoRoot 'CHANGELOG_V5.md') -Raw -Encoding UTF8
    if (-not $changelog.Contains("RuntimeVersion ``$version``")) { throw "CHANGELOG_V5.md does not document RuntimeVersion $version." }
    Write-Output ("Runtime version documented: {0}" -f $version)
}

function Assert-PackageContents([string]$OutputRoot) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $forbiddenFragments = @(
        'app/tools/selfcheck.py', 'app/tools/ci-selfcheck.ps1', 'app/tools/ci-requirements.txt', 'app/tools/fixtures/',
        'app/tools/scale-benchmark.ps1', 'app/tools/history-logic-selfcheck.ps1',
        'app/tools/create-pdf-diff-corpus.py', 'app/tools/pdf-diff-corpus-selfcheck.ps1',
        'tests/pdf-diff-corpus.mjs', 'docs/PDF_DIFF_CORPUS.md', 'docs/benchmarks/', '.github/'
    )
    $archives = @(Get-ChildItem -LiteralPath $OutputRoot -File -Filter '*.zip')
    if ($archives.Count -ne 2) { throw "Expected two ZIP releases, found $($archives.Count)." }
    foreach ($archive in $archives) {
        $zip = [IO.Compression.ZipFile]::OpenRead($archive.FullName)
        try {
            $entries = @($zip.Entries | ForEach-Object { $_.FullName.Replace('\','/') })
            $manifestEntries = @($entries | Where-Object { $_ -eq 'release-manifest.json' -or $_.EndsWith('/release-manifest.json', [StringComparison]::OrdinalIgnoreCase) })
            if ($manifestEntries.Count -ne 1) { throw "$($archive.Name) must contain exactly one release-manifest.json." }
            foreach ($fragment in $forbiddenFragments) {
                if (@($entries | Where-Object { $_.IndexOf($fragment, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count -gt 0) {
                    throw "$($archive.Name) contains development asset: $fragment"
                }
            }
        } finally { $zip.Dispose() }
    }
    $shared = @(Get-ChildItem -LiteralPath $OutputRoot -Directory)
    if ($shared.Count -ne 1) { throw "Expected one shared-folder release, found $($shared.Count)." }
    foreach ($fragment in $forbiddenFragments) {
        $relative = $fragment.Replace('/', '\').TrimEnd('\')
        if (Test-Path -LiteralPath (Join-Path $shared[0].FullName $relative)) { throw "Shared-folder release contains development asset: $fragment" }
    }
    Write-Output 'Release package content check OK.'
}

function Remove-CiGeneratedFiles {
    foreach ($path in @(
        (Join-Path $appRoot 'thirdparty-cache'),
        (Join-Path $appRoot 'lib\pdfbox\classes'),
        (Join-Path $repoRoot 'tmp\pdfs\pr7-final-composition'),
        (Join-Path $toolsRoot '__pycache__'),
        (Join-Path $repoRoot 'tests\__pycache__')
    )) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

$python = Get-PythonCommand
$node = (Get-Command node.exe -ErrorAction Stop).Source

try {
    Assert-PowerShellSyntax
    Assert-VersionDocumented
    Invoke-Checked 'Third-party verification' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'verify-thirdparty.ps1'),'-RequirePdfJs','-RequirePortableJava')
    Remove-CiGeneratedFiles
    Invoke-Checked 'Repository selfcheck' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'selfcheck.py')))
    Invoke-Checked 'JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\diff-regression.mjs'))
    Invoke-Checked 'PDF corpus JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\pdf-diff-corpus.mjs'))
    Invoke-Checked 'Diff regression' $node @((Join-Path $repoRoot 'tests\diff-regression.mjs'))
    Invoke-Checked 'Practical PDF diff corpus' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'pdf-diff-corpus-selfcheck.ps1'))
    Invoke-Checked 'Final composition regression' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'final-composition-selfcheck.py')))
    Invoke-Checked 'Operational readiness regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'operational-readiness-selfcheck.ps1'))
    Invoke-Checked 'Pack lifecycle regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'pack-lifecycle-selfcheck.ps1'))
    Invoke-Checked 'History logic regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'history-logic-selfcheck.ps1'))
    Invoke-Checked 'Custom pack workflow regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'custom-pack-workflow-selfcheck.ps1'))

    if (-not $SkipPackageSmoke) {
        New-Item -ItemType Directory -Path $PackageOutputDir -Force | Out-Null
        Invoke-Checked 'ZIP release smoke test' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'package-release.ps1'),'-OutputDir',$PackageOutputDir,'-NoOpen')
        Invoke-Checked 'Shared-folder release smoke test' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'package-release.ps1'),'-OutputDir',$PackageOutputDir,'-SharedFolderOnly','-NoOpen')
        Assert-PackageContents $PackageOutputDir
    }
    Write-Output 'ReportBinder CI selfcheck OK.'
} finally {
    Remove-CiGeneratedFiles
    if ($ownsPackageOutput -and (Test-Path -LiteralPath $PackageOutputDir)) {
        Remove-Item -LiteralPath $PackageOutputDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
