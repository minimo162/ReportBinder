# Scope=Fast は変更ごとの門。同梱JREもjarのビルドも要らない検査だけを並べる。
# Scope=Full は受け入れスイート全部。手動実行とリリース前だけ回す。
param(
    [ValidateSet('Full','Fast')]
    [string]$Scope = 'Full',
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

# 重い2本を同時に走らせる。どちらもCPUを使い切る実計算で、合計343秒のうち
# 短いほうがそのまま短縮できる。互いに参照するファイルは無く、作業先も別。
# 出力は各々ファイルへ取り、終わってからラベル順に出す。逐次実行のときと同じ
# 見え方にして、失敗の切り分けが変わらないようにするため。
function Invoke-CheckedInParallel([hashtable[]]$Checks) {
    $entries = @()
    foreach ($check in $Checks) {
        $stem = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('ci-' + [Guid]::NewGuid().ToString('N')))
        $entries += [pscustomobject]@{
            Label = [string]$check.Label
            LogPath = ($stem + '.log')
            CodePath = ($stem + '.code')
            Job = Start-Job -ScriptBlock {
                param($FilePath, $Arguments, $LogPath, $CodePath, $WorkingDirectory)
                # ネイティブの stderr で止まらないようにする(Invoke-NativeCapture と同じ理由)。
                $ErrorActionPreference = 'Continue'
                Set-Location $WorkingDirectory
                & $FilePath @Arguments *>&1 | Out-File -LiteralPath $LogPath -Encoding UTF8
                # 終了コードはファイルで受け渡す。ジョブの State から読み取るのは当てにならない。
                Set-Content -LiteralPath $CodePath -Value ([string]$LASTEXITCODE) -Encoding ASCII
            } -ArgumentList $check.FilePath, $check.Arguments, ($stem + '.log'), ($stem + '.code'), (Get-Location).Path
        }
    }
    $failures = @()
    foreach ($entry in $entries) {
        [void](Wait-Job -Job $entry.Job)
        Write-Output ("== {0} ==" -f $entry.Label)
        if (Test-Path -LiteralPath $entry.LogPath) {
            Get-Content -LiteralPath $entry.LogPath -ErrorAction SilentlyContinue | ForEach-Object { Write-Output $_ }
        }
        $code = $null
        if (Test-Path -LiteralPath $entry.CodePath) {
            $code = [string](Get-Content -LiteralPath $entry.CodePath -Raw -ErrorAction SilentlyContinue).Trim()
        }
        Receive-Job -Job $entry.Job -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $entry.LogPath, $entry.CodePath -Force -ErrorAction SilentlyContinue
        # 終了コードが取れない場合も失敗として扱う(ジョブが落ちて書けなかった場合)。
        if ([string]::IsNullOrWhiteSpace($code) -or $code -ne '0') {
            $shown = $code
            if ([string]::IsNullOrWhiteSpace($shown)) { $shown = 'unknown' }
            $failures += ("{0} (exit {1})" -f $entry.Label, $shown)
        }
    }
    if ($failures.Count -gt 0) { throw (($failures -join '; ') + ' failed.') }
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

# 配布物に入るファイルが変わったのに版番号が据え置きだと、既にその版を持つPCは
# 共有フォルダーを読まずローカルコピーで起動し続ける。launch.ps1 の「ローカル版は
# 最新か」の判定は installed.json / server.ps1 / web\app.js の**存在**だけを見て
# おり、内容も日付も比べない。実際に #59 #60 の修正が届かない状態になっていた。
#
# 除外の判断は package-release.ps1 の除外一覧をその場で読んで行う。ここに写しを
# 置くと、片方だけ直されたときに黙って判定がずれる。
function Get-ReleaseExcludedPaths {
    $packager = Join-Path $toolsRoot 'package-release.ps1'
    $source = Get-Content -LiteralPath $packager -Raw -Encoding UTF8
    $start = $source.IndexOf('function Remove-ReleaseDevelopmentFiles', [StringComparison]::Ordinal)
    if ($start -lt 0) { throw 'package-release.ps1 に Remove-ReleaseDevelopmentFiles が見つかりません。除外一覧を読めないため版番号の門を判定できません。' }
    # 終端は「次の関数定義」で切る。`)) {` のような書式そのものを目印にすると、
    # 整形を1文字変えただけで窓が次の関数まで伸び、無関係な文字列を除外一覧として
    # 拾ってしまう（'app' を拾うと全変更が除外され、門は緑のまま素通りする）。
    $end = $source.IndexOf("`nfunction ", $start + 1, [StringComparison]::Ordinal)
    if ($end -lt 0) { $end = $source.Length }
    $paths = @([regex]::Matches($source.Substring($start, $end - $start), "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    # 読み取りに失敗したまま「除外0件」で進むと全変更が配布対象扱いになり、逆に
    # 一覧を広く拾いすぎると全変更が除外されて門が素通りする。どちらも黙って
    # 起きるため、一覧が想定の形をしていることを錨で確かめてから使う。
    foreach ($anchor in @('app\tools\selfcheck.py', 'app\tools\package-release.ps1', '.github')) {
        if ($paths -notcontains $anchor) { throw "package-release.ps1 の除外一覧を正しく読み取れませんでした（$anchor が見つかりません／$($paths.Count) 件）。" }
    }
    # 読み取り窓が次の関数まで伸びて 'app' のような語を拾うと、app/ 配下が丸ごと
    # 除外され、門は「配布対象の変更なし」と言って緑のまま素通りする。錨だけでは
    # 拾いすぎを検知できない（錨も同じ一覧の中にあるため）ので、門を無力化する
    # 項目そのものを名指しで拒む。tests のような正当な最上位ディレクトリは通す。
    foreach ($path in $paths) {
        if ($path.TrimEnd('\') -eq 'app') { throw 'package-release.ps1 の除外一覧が app 全体を指しています。読み取り範囲がずれている可能性があります。' }
    }
    return $paths
}

# 版番号は 2026.08.15.1 形式。数値の組にして大小を比べる。文字列比較だと
# 2026.08.09.10 < 2026.08.09.9 になり、連番が2桁へ乗った日に判定が反転する。
function ConvertTo-RuntimeVersionTuple([string]$Version) {
    $m = [regex]::Match($Version, '^(\d{4})\.(\d{2})\.(\d{2})\.(\d+)$')
    if (-not $m.Success) { return $null }
    return @([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value, [int]$m.Groups[4].Value)
}

function Compare-RuntimeVersion($Left, $Right) {
    for ($i = 0; $i -lt 4; $i++) {
        if ($Left[$i] -ne $Right[$i]) { return ($Left[$i] - $Right[$i]) }
    }
    return 0
}

function Assert-RuntimeVersionBumped {
    $baseRef = 'origin/main'
    if (-not [string]::IsNullOrWhiteSpace($env:REPORTBINDER_BASE_REF)) { $baseRef = [string]$env:REPORTBINDER_BASE_REF }
    elseif (-not [string]::IsNullOrWhiteSpace($env:GITHUB_BASE_REF)) { $baseRef = 'origin/' + [string]$env:GITHUB_BASE_REF }

    # native コマンドの stderr を 2>&1 で成功ストリームへ流すと、PowerShell 5.1 は
    # 各行を ErrorRecord にする。$ErrorActionPreference='Stop' の下ではそれが終端
    # エラーになり、下の SKIPPED 分岐へ一度も到達しない（.git の無いコピー、
    # dubious ownership、git が PATH に無いときに必ず起きる）。混ぜないこと。
    & git -C $repoRoot rev-parse --verify --quiet ($baseRef + '^{commit}') | Out-Null
    if ($LASTEXITCODE -ne 0) {
        # 解決できないまま素通りさせると門が無いのと同じになる。飛ばしたことを必ず出す。
        Write-Output ("SKIPPED: runtime version bump gate ({0} を解決できません。履歴の取得設定を確認してください)" -f $baseRef)
        return
    }

    # 三点は「基準から分岐したあとの変更」。squash マージ運用では、先行PRがマージ
    # された時点で後続PRの merge-base が分岐元まで戻り、先行PRの変更が混ざる。
    # 二点との積を取ると、基準側に既に同じ内容が入っているファイルが落ちる。
    $threeDot = @(& git -C $repoRoot diff --name-only ($baseRef + '...HEAD') -- 'app/')
    if ($LASTEXITCODE -ne 0) { throw '版番号の門: git diff (three-dot) に失敗しました。' }
    $twoDot = @(& git -C $repoRoot diff --name-only $baseRef 'HEAD' -- 'app/')
    if ($LASTEXITCODE -ne 0) { throw '版番号の門: git diff (two-dot) に失敗しました。' }
    $changed = @($threeDot | Where-Object { $twoDot -contains $_ })

    if ($changed.Count -eq 0) {
        # app/ を触っていない変更まで、除外一覧の読み取り失敗で巻き添えにしない。
        Write-Output ("Runtime version gate: app/ の変更なし（基準 {0}）" -f $baseRef)
        return
    }

    $excluded = Get-ReleaseExcludedPaths
    $shipping = @()
    foreach ($entry in $changed) {
        $rel = ([string]$entry).Trim()
        if ($rel -eq '' -or $rel -eq 'app/runtime-version.json') { continue }
        $win = $rel.Replace('/', '\')
        $isExcluded = $false
        foreach ($ex in $excluded) {
            if ($win -eq $ex -or $win.StartsWith($ex + '\', [StringComparison]::OrdinalIgnoreCase)) { $isExcluded = $true; break }
        }
        if (-not $isExcluded) { $shipping += $rel }
    }
    if ($shipping.Count -eq 0) {
        Write-Output ("Runtime version gate: 配布対象の変更なし（基準 {0}）" -f $baseRef)
        return
    }

    $baseJson = (& git -C $repoRoot show ($baseRef + ':app/runtime-version.json'))
    if ($LASTEXITCODE -ne 0) {
        Write-Output ("Runtime version gate: 基準 {0} に runtime-version.json がないため比較を省略します" -f $baseRef)
        return
    }
    # HEAD 側もコミット済みの内容から読む。片方を作業ツリーから読むと、ローカルで
    # 緑になった判定が push した途端に CI で落ちる（逆も起きる）。
    $headJson = (& git -C $repoRoot show 'HEAD:app/runtime-version.json')
    if ($LASTEXITCODE -ne 0) { throw '版番号の門: HEAD の runtime-version.json を読めませんでした。' }
    $baseVersion = ([string](($baseJson -join "`n" | ConvertFrom-Json).version)).Trim()
    $headVersion = ([string](($headJson -join "`n" | ConvertFrom-Json).version)).Trim()

    $shown = ($shipping | Select-Object -First 10) -join ', '
    if ($shipping.Count -gt 10) { $shown += (' ほか {0} 件' -f ($shipping.Count - 10)) }
    $howTo = "app\runtime-version.json を上げ、CHANGELOG_V5.md に RuntimeVersion ``<新版>`` を書いてください。" +
             "先行PRのマージ後に出た場合は、基準ブランチを取り込み直す(rebase)と消えることがあります。" +
             ("`n対象: {0}" -f $shown)

    $baseTuple = ConvertTo-RuntimeVersionTuple $baseVersion
    $headTuple = ConvertTo-RuntimeVersionTuple $headVersion
    if ($null -eq $baseTuple -or $null -eq $headTuple) {
        throw ("版番号の門: 版番号を数値として読めません（基準 {0} / HEAD {1}）。" -f $baseVersion, $headVersion)
    }
    # 「変わっていれば通す」だと、巻き戻し(過去の版へ戻す)が素通りする。既にその版を
    # 起動したPCには同名のローカルコピーがあり、launch.ps1 はそれを使い続けるため、
    # 巻き戻しは据え置きと同じ結果になる。真に新しいことを要求する。
    if ((Compare-RuntimeVersion $headTuple $baseTuple) -le 0) {
        throw ("配布対象のファイルが変わりましたが、RuntimeVersion {0} は基準の {1} より新しくありません。" -f $headVersion, $baseVersion) +
              "既にその版を起動したPCにはローカルコピーが残っており、共有フォルダーを読まないため変更は届きません。" + $howTo
    }
    # 同じ日に2つのPRが同じ連番を選ぶと、両方ともこの門を通り、後からマージされた
    # 側の変更が永久に届かなくなる（#59 #60 と同じ結末）。基準時点の CHANGELOG に
    # 既に書かれている版番号は、他のPRが使い終えたものとみなして拒む。
    $baseChangelog = (& git -C $repoRoot show ($baseRef + ':CHANGELOG_V5.md'))
    if ($LASTEXITCODE -eq 0) {
        $needle = 'RuntimeVersion `' + $headVersion + '`'
        if (($baseChangelog -join "`n").Contains($needle)) {
            throw ("RuntimeVersion {0} は基準 {1} の CHANGELOG_V5.md に既にあります。" -f $headVersion, $baseRef) +
                  "同じ版番号を2度使うと、先に配った側を起動したPCへ後の変更が届きません。次の連番を使ってください。" + $howTo
        }
    }
    Write-Output ("Runtime version gate: {0} -> {1}（配布対象 {2} 件）" -f $baseVersion, $headVersion, $shipping.Count)
}

function Assert-PackageContents([string]$OutputRoot) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $forbiddenFragments = @(
        'app/tools/selfcheck.py', 'app/tools/ci-selfcheck.ps1', 'app/tools/ci-requirements.txt', 'app/tools/fixtures/',
        'app/tools/scale-benchmark.ps1', 'app/tools/history-logic-selfcheck.ps1',
        'app/tools/create-pdf-diff-corpus.py', 'app/tools/pdf-diff-corpus-selfcheck.ps1',
        'app/tools/pack-lifecycle-selfcheck.ps1', 'app/tools/custom-pack-workflow-selfcheck.ps1',
        'tests/',
        'docs/PDF_DIFF_CORPUS.md', 'docs/benchmarks/', '.github/'
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
    Assert-RuntimeVersionBumped

    if ($Scope -eq 'Fast') {
        Remove-CiGeneratedFiles
        Invoke-Checked 'Repository selfcheck (static)' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'selfcheck.py'),'--static-only'))
        # structure.json のスキーマ移行だけは変更ごとに見る。利用者データを不可逆に
        # 壊しうる唯一の種類で、しかも外部依存ゼロの1.7秒で確かめられるため。
        Invoke-Checked 'Schema migration regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'schema-v3-selfcheck.ps1'))
        Invoke-Checked 'JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\diff-regression.mjs'))
        Invoke-Checked 'PDF corpus JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\pdf-diff-corpus.mjs'))
        # 何を見ていないかを結果に残す。通った表示だけが後から参照されるため。
        Write-Output 'ReportBinder CI fast gate OK. (regression suites and package smoke tests were not run)'
        return
    }

    Invoke-Checked 'Third-party verification' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'verify-thirdparty.ps1'),'-RequirePdfJs','-RequirePortableJava')
    Remove-CiGeneratedFiles
    Invoke-Checked 'Repository selfcheck' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'selfcheck.py')))
    Invoke-Checked 'JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\diff-regression.mjs'))
    Invoke-Checked 'PDF corpus JavaScript syntax' $node @('--check',(Join-Path $repoRoot 'tests\pdf-diff-corpus.mjs'))
    Invoke-CheckedInParallel @(
        @{ Label = 'Diff regression'; FilePath = $node; Arguments = @((Join-Path $repoRoot 'tests\diff-regression.mjs')) },
        @{ Label = 'Practical PDF diff corpus'; FilePath = 'powershell.exe'; Arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'pdf-diff-corpus-selfcheck.ps1')) }
    )
    Invoke-Checked 'Final composition regression' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'final-composition-selfcheck.py')))
    Invoke-Checked 'Operational readiness regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'operational-readiness-selfcheck.ps1'))
    Invoke-Checked 'Pack lifecycle regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'pack-lifecycle-selfcheck.ps1'))
    Invoke-Checked 'History logic regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'history-logic-selfcheck.ps1'))
    Invoke-Checked 'Custom pack workflow regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'custom-pack-workflow-selfcheck.ps1'))
    # この3本は書かれていたのに、どのCI経路からも呼ばれていなかった。原稿アダプタは
    # V5で汎用化した経路そのもので、回帰網に穴が空いたままだった。
    Invoke-Checked 'PDF source adapter regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'pdf-source-adapter-selfcheck.ps1'))
    Invoke-Checked 'History generalization regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'history-generalization-selfcheck.ps1'))
    # Word のアダプタ検証だけは DOCX の実物と、実際の Word による変換が要る
    # (他のアダプタ検証と違い、変換を差し替えていない)。Office の入っている開発機では
    # 実行し、入っていないCIランナーでは飛ばす。飛ばしたことは必ず出力に残す。
    if ($null -eq [Type]::GetTypeFromProgID('Word.Application')) {
        Write-Output '== Word source adapter regression =='
        Write-Output 'SKIPPED: Word is not installed on this machine, so the real conversion path cannot run here.'
    } else {
        $wordFixtureDir = Join-Path ([IO.Path]::GetTempPath()) ('ReportBinderWordFixtures_' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $wordFixtureDir -Force | Out-Null
        try {
            Invoke-Checked 'Word adapter fixtures' $python.file (@($python.prefix) + @((Join-Path $toolsRoot 'create-word-adapter-fixtures.py'),'--output-dir',$wordFixtureDir))
            Invoke-Checked 'Word source adapter regression' 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $toolsRoot 'word-source-adapter-selfcheck.ps1'),'-FixtureV1',(Join-Path $wordFixtureDir 'word-fixture-v1.docx'),'-FixtureV2',(Join-Path $wordFixtureDir 'word-fixture-v2.docx'))
        } finally {
            Remove-Item -LiteralPath $wordFixtureDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

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
