param(
    [ValidateSet('medium','large')]
    [string]$Tier = 'large',
    [string]$PythonPath = 'python',
    [string]$NodePath = 'node',
    [Parameter(Mandatory=$true)]
    [string]$NodeModulesPath,
    [string]$ReportPath = '',
    [switch]$KeepWorkspace
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$repoRoot = Split-Path -Parent $appRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$pythonFixturePath = Join-Path $toolsRoot 'create-scale-benchmark-fixtures.py'
$workbookFixturePath = Join-Path $toolsRoot 'create-scale-benchmark-workbooks.mjs'
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('ReportBinderScale_' + [Guid]::NewGuid().ToString('N'))
$submissionDir = Join-Path $workRoot 'submission'
$configRoot = Join-Path $workRoot 'local-config'
$runnerPath = Join-Path $workRoot 'create-scale-benchmark-workbooks.mjs'
$nodeModulesJunction = Join-Path $workRoot 'node_modules'
$stdoutPath = Join-Path $workRoot 'server.stdout.log'
$stderrPath = Join-Path $workRoot 'server.stderr.log'
$serverProcess = $null
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')

$tierConfig = if ($Tier -eq 'large') {
    [ordered]@{ pdfCount=6; pdfPages=20; wordCount=3; wordPages=10; excelCount=3; excelSheets=8; excelRows=45 }
} else {
    [ordered]@{ pdfCount=4; pdfPages=12; wordCount=2; wordPages=6; excelCount=2; excelSheets=5; excelRows=32 }
}

function Get-DirectoryBytes([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return [int64]0 }
    $sum = [int64]0
    Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { $sum += [int64]$_.Length } catch { }
    }
    return $sum
}

function Get-SourceBytes([string]$Path) {
    $sum = [int64]0
    Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop | Where-Object {
        $_.Extension.ToLowerInvariant() -in @('.pdf','.xlsx','.docx')
    } | ForEach-Object { $sum += [int64]$_.Length }
    return $sum
}

function Get-FreeBytes([string]$Path) {
    $root = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Path))
    return [int64]([IO.DriveInfo]::new($root).AvailableFreeSpace)
}

function Get-FreePort {
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    try {
        $listener.Start()
        return [int]$listener.LocalEndpoint.Port
    } finally {
        $listener.Stop()
    }
}

function New-Token {
    $bytes = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return ([Convert]::ToBase64String($bytes)).TrimEnd('=').Replace('+','-').Replace('/','_')
}

function Invoke-LocalApi([string]$Method, [string]$Path, $Body = $null) {
    $uri = "http://127.0.0.1:$script:port$Path"
    $params = @{ Uri=$uri; Method=$Method; Headers=@{'X-ReportBinder-Token'=$script:token}; TimeoutSec=900 }
    if ($null -ne $Body) {
        $params.ContentType = 'application/json; charset=utf-8'
        $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress)
    }
    return Invoke-RestMethod @params
}

function Wait-ServerReady([int]$TimeoutSeconds = 45) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        if ($script:serverProcess.HasExited) {
            $stderr = if (Test-Path -LiteralPath $script:stderrPath) { Get-Content -LiteralPath $script:stderrPath -Raw -ErrorAction SilentlyContinue } else { '' }
            throw "ReportBinder server stopped before becoming ready. $stderr"
        }
        try { [void](Invoke-LocalApi 'GET' '/api/state'); return } catch { Start-Sleep -Milliseconds 300 }
    } while ([DateTime]::UtcNow -lt $deadline)
    throw 'ReportBinder server readiness timed out.'
}

function Wait-RenderJob([string]$JobId, [int]$TimeoutSeconds = 1200) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $peakBytes = Get-DirectoryBytes $script:workRoot
    do {
        $job = Invoke-LocalApi 'POST' '/api/jobs/status' @{ jobId=$JobId }
        $currentBytes = Get-DirectoryBytes $script:workRoot
        if ($currentBytes -gt $peakBytes) { $peakBytes = $currentBytes }
        $status = [string]$job.status
        if ($status -in @('completed','completed-with-errors','failed')) {
            return [pscustomobject][ordered]@{ job=$job; peakWorkspaceBytes=$peakBytes }
        }
        Start-Sleep -Milliseconds 700
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Render job timed out: $JobId"
}

function Add-Phase([Collections.ArrayList]$Phases, [string]$Name, [Diagnostics.Stopwatch]$Timer, [int64]$PeakBytes = 0, $Details = $null) {
    $Timer.Stop()
    if ($PeakBytes -le 0) { $PeakBytes = Get-DirectoryBytes $script:workRoot }
    [void]$Phases.Add([ordered]@{
        name=$Name
        durationSeconds=[Math]::Round($Timer.Elapsed.TotalSeconds, 3)
        workspaceBytes=(Get-DirectoryBytes $script:workRoot)
        peakObservedWorkspaceBytes=$PeakBytes
        freeBytes=(Get-FreeBytes $script:workRoot)
        details=$Details
    })
}

function Assert-RenderSucceeded($Job, [string]$Phase) {
    if ([string]$Job.status -ne 'completed' -or [int]$Job.failed -gt 0) {
        throw "$Phase failed: $($Job | ConvertTo-Json -Depth 12 -Compress)"
    }
}

function ConvertTo-ReadinessSummary($Volume) {
    $outputPath = [string]$Volume.outputPdf
    return [ordered]@{
        canBuild=[bool]$Volume.canBuild
        pageCount=[int]$Volume.pageCount
        status=[string]$Volume.status
        displayState=[string]$Volume.displayState
        builtFingerprint=[string]$Volume.builtFingerprint
        currentFingerprint=[string]$Volume.currentFingerprint
        outputPdfName=if ([string]::IsNullOrWhiteSpace($outputPath)) { '' } else { [IO.Path]::GetFileName($outputPath) }
        outputPdfExists=[bool]$Volume.outputPdfExists
        lastBuiltAt=[string]$Volume.lastBuiltAt
        blockers=@($Volume.blockers | ForEach-Object { [string]$_ })
        staleReasons=@($Volume.staleReasons | ForEach-Object { [string]$_ })
    }
}

function Get-Phase([Collections.ArrayList]$Phases, [string]$Name) {
    return @($Phases | Where-Object { [string]$_.name -eq $Name })[0]
}

function Remove-BenchmarkWorkspace([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if (-not $full.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or -not ([IO.Path]::GetFileName($full)).StartsWith('ReportBinderScale_')) {
        throw "Unexpected benchmark cleanup target: $full"
    }
    if (Test-Path -LiteralPath $script:nodeModulesJunction) {
        $junction = Get-Item -LiteralPath $script:nodeModulesJunction -Force
        if (-not ($junction.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Benchmark node_modules is not a reparse point.' }
        [IO.Directory]::Delete($script:nodeModulesJunction)
    }
    if ([IO.Directory]::Exists($full)) { [IO.Directory]::Delete('\\?\' + $full, $true) }
}

$phases = New-Object Collections.ArrayList
$report = $null

try {
    foreach ($directory in @($workRoot,$submissionDir,$configRoot)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $PythonPath -PathType Leaf) -and -not (Get-Command $PythonPath -ErrorAction SilentlyContinue)) { throw "Python not found: $PythonPath" }
    if (-not (Test-Path -LiteralPath $NodePath -PathType Leaf) -and -not (Get-Command $NodePath -ErrorAction SilentlyContinue)) { throw "Node.js not found: $NodePath" }
    $NodeModulesPath = [IO.Path]::GetFullPath($NodeModulesPath)
    if (-not (Test-Path -LiteralPath (Join-Path $NodeModulesPath '@oai\artifact-tool') -PathType Container)) { throw "@oai/artifact-tool not found under: $NodeModulesPath" }
    Copy-Item -LiteralPath $workbookFixturePath -Destination $runnerPath -Force
    New-Item -ItemType Junction -Path $nodeModulesJunction -Target $NodeModulesPath | Out-Null

    $timer = [Diagnostics.Stopwatch]::StartNew()
    & $PythonPath $pythonFixturePath --output $submissionDir --pdf-count $tierConfig.pdfCount --pdf-pages $tierConfig.pdfPages --word-count $tierConfig.wordCount --word-pages $tierConfig.wordPages --version 1
    if ($LASTEXITCODE -ne 0) { throw 'PDF/Word scale fixture generation failed.' }
    & $NodePath $runnerPath --output $submissionDir --count $tierConfig.excelCount --sheets $tierConfig.excelSheets --rows $tierConfig.excelRows
    $createdExcelFiles = @(Get-ChildItem -LiteralPath $submissionDir -Filter 'large-excel-*.xlsx' -File)
    if ($createdExcelFiles.Count -ne [int]$tierConfig.excelCount) { throw 'Excel scale fixture generation failed.' }
    $sourceFiles = @(Get-ChildItem -LiteralPath $submissionDir -File | Where-Object { $_.Extension.ToLowerInvariant() -in @('.pdf','.xlsx','.docx') })
    $sourceBytes = Get-SourceBytes $submissionDir
    Add-Phase $phases 'fixture-generation' $timer 0 ([ordered]@{ sourceCount=$sourceFiles.Count; sourceBytes=$sourceBytes })

    $script:port = Get-FreePort
    $script:token = New-Token
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'powershell.exe'
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$serverPath`" -Mode ja -Port $port -Token $token -NoOpen"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $script:serverProcess = [Diagnostics.Process]::new()
    $script:serverProcess.StartInfo = $startInfo
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $configRoot, 'Process')
    try {
        if (-not $script:serverProcess.Start()) { throw 'ReportBinder server process could not be started.' }
    } finally {
        [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    }
    Wait-ServerReady

    $diagnosticsResponse = Invoke-LocalApi 'POST' '/api/diagnostics/run' @{}
    $diagnostics = $diagnosticsResponse.diagnostics
    if (-not [bool]$diagnostics.office.excel.available -or -not [bool]$diagnostics.office.word.available) {
        throw "Interactive Office session is required (Excel=$($diagnostics.office.excel.status), Word=$($diagnostics.office.word.status)). Run the benchmark from the signed-in Windows desktop session."
    }
    $pathsResponse = Invoke-LocalApi 'POST' '/api/paths' @{ submissionDir=$submissionDir }
    $paths = $pathsResponse.paths
    [void](Invoke-LocalApi 'PATCH' '/api/v2/packs/pack_ecm' @{
        documentTitle="Scale benchmark $Tier"
        documentSubtitle='Synthetic mixed-source acceptance workload'
        outputFileNamePattern='{projectId}_{packName}_{targetName}_{yyyyMMdd}.pdf'
    })

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $candidates = Invoke-LocalApi 'GET' '/api/submission-files'
    $relativePaths = @($candidates.files | ForEach-Object { [string]$_.relativePath })
    $registered = Invoke-LocalApi 'POST' '/api/workbooks/register-batch' @{ category='ecm'; relativePaths=$relativePaths }
    $sourceIds = @($registered.state.structure.workbooks | Where-Object { [string]$_.category -eq 'ecm' } | ForEach-Object { [string]$_.workbookId })
    Add-Phase $phases 'registration' $timer 0 ([ordered]@{ candidateCount=$relativePaths.Count; registeredCount=$sourceIds.Count })

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $started = Invoke-LocalApi 'POST' '/api/workbooks/render/start' @{ category='ecm'; workbookIds=$sourceIds }
    $render = Wait-RenderJob ([string]$started.jobId)
    Assert-RenderSucceeded $render.job 'Initial mixed-source render'
    Add-Phase $phases 'initial-render' $timer ([int64]$render.peakWorkspaceBytes) ([ordered]@{ total=$render.job.total; completed=$render.job.completed; failed=$render.job.failed })

    $state = Invoke-LocalApi 'GET' '/api/state'
    $sourceIdLookup = @{}
    foreach ($sourceId in $sourceIds) { $sourceIdLookup[[string]$sourceId] = $true }
    $pages = @($state.structure.pages | Where-Object { $sourceIdLookup.ContainsKey([string]$_.workbookId) })
    if ($pages.Count -lt 1) { throw 'No pages were produced by the scale render.' }
    $mainLogicalCount = [Math]::Ceiling($pages.Count * 0.8)
    $mainIds = @($pages | Select-Object -First $mainLogicalCount | ForEach-Object { [string]$_.pageId })
    $appendixIds = @($pages | Select-Object -Skip $mainLogicalCount | ForEach-Object { [string]$_.pageId })
    $timer = [Diagnostics.Stopwatch]::StartNew()
    [void](Invoke-LocalApi 'POST' '/api/pages/reorder' @{ category='ecm'; volumes=@{'ja-main'=$mainIds; 'ja-appendix'=$appendixIds; none=@()} })
    $buildInitial = Invoke-LocalApi 'POST' '/api/final/build-all' @{ category='ecm' }
    Add-Phase $phases 'initial-final-build' $timer 0 ([ordered]@{ logicalPages=$pages.Count; mainLogicalPages=$mainIds.Count; appendixLogicalPages=$appendixIds.Count })

    $timer = [Diagnostics.Stopwatch]::StartNew()
    & $PythonPath $pythonFixturePath --output $submissionDir --pdf-count $tierConfig.pdfCount --pdf-pages $tierConfig.pdfPages --word-count $tierConfig.wordCount --word-pages $tierConfig.wordPages --version 2 --update-only
    if ($LASTEXITCODE -ne 0) { throw 'PDF/Word fixture update failed.' }
    & $NodePath $runnerPath --output $submissionDir --count $tierConfig.excelCount --sheets $tierConfig.excelSheets --rows $tierConfig.excelRows --update-only
    if (-not (Test-Path -LiteralPath (Join-Path $submissionDir 'large-excel-01-v2-preview.png') -PathType Leaf)) { throw 'Excel fixture update failed.' }
    $scan = Invoke-LocalApi 'POST' '/api/scan-updates' @{ forceHash=$true }
    $updatedState = Invoke-LocalApi 'GET' '/api/state'
    $updatedSources = @($updatedState.structure.workbooks | Where-Object { [string]$_.status -in @('new','excel-updated','source-updated','render-error') })
    Add-Phase $phases 'update-detection' $timer 0 ([ordered]@{ updatedSourceCount=$updatedSources.Count; capturedSnapshotCount=@($scan.capturedSnapshots).Count })

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $startedUpdate = Invoke-LocalApi 'POST' '/api/workbooks/render/start' @{ category='ecm'; onlyUpdated=$true }
    $renderUpdate = Wait-RenderJob ([string]$startedUpdate.jobId)
    Assert-RenderSucceeded $renderUpdate.job 'Updated-source render'
    Add-Phase $phases 'updated-render-and-analysis' $timer ([int64]$renderUpdate.peakWorkspaceBytes) ([ordered]@{ total=$renderUpdate.job.total; completed=$renderUpdate.job.completed; failed=$renderUpdate.job.failed })

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $buildUpdated = Invoke-LocalApi 'POST' '/api/final/build-all' @{ category='ecm' }
    Add-Phase $phases 'updated-final-build' $timer

    $finalState = Invoke-LocalApi 'GET' '/api/state'
    $readiness = Invoke-LocalApi 'GET' '/api/final/readiness?category=ecm'
    $outputFiles = @(Get-ChildItem -LiteralPath ([string]$paths.outputDir) -Filter '*.pdf' -File -ErrorAction SilentlyContinue)
    $dataBytes = Get-DirectoryBytes ([string]$paths.dataDir)
    $outputBytes = Get-DirectoryBytes ([string]$paths.outputDir)
    $peakWorkspaceBytes = [int64](($phases | ForEach-Object { [int64]$_.peakObservedWorkspaceBytes } | Measure-Object -Maximum).Maximum)
    $mainSummary = ConvertTo-ReadinessSummary $readiness.volumes.'ja-main'
    $appendixSummary = ConvertTo-ReadinessSummary $readiness.volumes.'ja-appendix'
    $finalOutputPageCount = [int]$mainSummary.pageCount + [int]$appendixSummary.pageCount
    $minimumFreeBytes = [int64](($phases | ForEach-Object { [int64]$_.freeBytes } | Measure-Object -Minimum).Minimum)
    $acceptanceThresholds = [ordered]@{
        initialRenderSeconds=300
        updateDetectionSeconds=60
        updatedRenderAndAnalysisSeconds=120
        finalBuildSecondsPerRun=60
        peakWorkspaceBytes=1GB
        minimumFreeBytes=10GB
    }
    $acceptanceChecks = [ordered]@{
        allSourcesRegistered=([int](Get-Phase $phases 'registration').details.registeredCount -eq $sourceFiles.Count)
        initialRenderSucceeded=([int](Get-Phase $phases 'initial-render').details.failed -eq 0)
        exactlyThreeSourcesDetected=($updatedSources.Count -eq 3)
        updatedRenderSucceeded=([int](Get-Phase $phases 'updated-render-and-analysis').details.failed -eq 0)
        bothOutputsBuilt=([bool]$mainSummary.outputPdfExists -and [bool]$appendixSummary.outputPdfExists)
        initialRenderWithinBudget=([double](Get-Phase $phases 'initial-render').durationSeconds -le $acceptanceThresholds.initialRenderSeconds)
        updateDetectionWithinBudget=([double](Get-Phase $phases 'update-detection').durationSeconds -le $acceptanceThresholds.updateDetectionSeconds)
        updatedRenderWithinBudget=([double](Get-Phase $phases 'updated-render-and-analysis').durationSeconds -le $acceptanceThresholds.updatedRenderAndAnalysisSeconds)
        initialBuildWithinBudget=([double](Get-Phase $phases 'initial-final-build').durationSeconds -le $acceptanceThresholds.finalBuildSecondsPerRun)
        updatedBuildWithinBudget=([double](Get-Phase $phases 'updated-final-build').durationSeconds -le $acceptanceThresholds.finalBuildSecondsPerRun)
        peakWorkspaceWithinBudget=($peakWorkspaceBytes -le $acceptanceThresholds.peakWorkspaceBytes)
        freeSpaceWithinBudget=($minimumFreeBytes -ge $acceptanceThresholds.minimumFreeBytes)
    }
    $acceptancePassed = (@($acceptanceChecks.Values | Where-Object { -not [bool]$_ }).Count -eq 0)
    $report = [ordered]@{
        schemaVersion=2
        benchmark='ReportBinder mixed-source scale acceptance'
        tier=$Tier
        measuredAt=(Get-Date).ToString('o')
        workspaceRetainedForQa=[bool]$KeepWorkspace
        environment=[ordered]@{
            status=$diagnostics.status
            excel=$diagnostics.office.excel
            word=$diagnostics.office.word
            java=$diagnostics.runtime.java
            pdfbox=$diagnostics.runtime.pdfbox
            pdfjs=$diagnostics.runtime.pdfjs
        }
        workload=[ordered]@{
            sourceCount=$sourceFiles.Count
            sourceBytes=$sourceBytes
            pdfCount=$tierConfig.pdfCount
            pdfPagesEach=$tierConfig.pdfPages
            wordCount=$tierConfig.wordCount
            wordPagesEach=$tierConfig.wordPages
            excelCount=$tierConfig.excelCount
            excelSheetsEach=$tierConfig.excelSheets
            excelRowsEach=$tierConfig.excelRows
            sourceLogicalPageCount=$pages.Count
            updatedSourceCount=$updatedSources.Count
        }
        result=[ordered]@{
            main=$mainSummary
            appendix=$appendixSummary
            finalOutputPageCount=$finalOutputPageCount
            generatedCompositionPageCount=($finalOutputPageCount - $pages.Count)
            outputFiles=@($outputFiles | ForEach-Object { [ordered]@{ name=$_.Name; bytes=[int64]$_.Length } })
            dataBytes=$dataBytes
            outputBytes=$outputBytes
            peakObservedWorkspaceBytes=$peakWorkspaceBytes
            minimumObservedFreeBytes=$minimumFreeBytes
            workspaceToSourceRatio=if($sourceBytes -gt 0){[Math]::Round($peakWorkspaceBytes / $sourceBytes, 2)}else{$null}
        }
        acceptance=[ordered]@{ passed=$acceptancePassed; thresholds=$acceptanceThresholds; checks=$acceptanceChecks }
        phases=@($phases)
    }

    $json = $report | ConvertTo-Json -Depth 30
    if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = [IO.Path]::GetFullPath($ReportPath)
        $parent = Split-Path -Parent $ReportPath
        if (-not [string]::IsNullOrWhiteSpace($parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        [IO.File]::WriteAllText($ReportPath, $json, [Text.UTF8Encoding]::new($false))
    }
    Write-Output $json
} catch {
    $line = [int]$_.InvocationInfo.ScriptLineNumber
    throw "Scale benchmark failed at line ${line}: $($_.Exception.Message)"
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    if ($serverProcess -and -not $serverProcess.HasExited) { Stop-Process -Id $serverProcess.Id -Force -ErrorAction SilentlyContinue }
    if (-not $KeepWorkspace) {
        try { Remove-BenchmarkWorkspace $workRoot } catch { Write-Warning $_.Exception.Message }
    } else {
        Write-Warning "Benchmark workspace retained: $workRoot"
    }
}
