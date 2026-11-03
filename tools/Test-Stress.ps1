param(
    [switch]$SkipBuild,
    [switch]$SkipIntegration,
    [switch]$SkipOwnership,
    [ValidateRange(0, 32)][int]$WorkerCount = 0,
    [ValidateRange(300, 5000)][int]$DiagnosticPingTimeoutMs = 300,
    [string]$MappingName = 'Local\UbeCPUStressIntegration',
    [string]$AppExecutable = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$testSources = Join-Path $root 'tests\stress'
$outputDirectory = Join-Path $root 'build\stress-tests'
[void][IO.Directory]::CreateDirectory($outputDirectory)
if (-not $SkipBuild) {
    # VB6 parses its source/project files as ANSI with CRLF line endings.
    foreach ($file in Get-ChildItem -LiteralPath $testSources -File | Where-Object { $_.Extension -in '.bas', '.vbp' }) {
        $contents = [IO.File]::ReadAllText($file.FullName)
        [IO.File]::WriteAllText($file.FullName, [regex]::Replace($contents, '\r?\n', "`r`n"), [Text.Encoding]::GetEncoding(1252))
    }
    & (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project 'tests\stress\StressTests.vbp' -OutputDirectory 'build\stress-tests'
    & (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project 'tests\stress\WorkerTests.vbp' -OutputDirectory 'build\stress-tests'
    if (-not $SkipIntegration) {
        & (Join-Path $PSScriptRoot 'Build-Stress.ps1') -TestIsolation
    }
}

$unitExe = Join-Path $outputDirectory 'StressTests.exe'
$unitLog = Join-Path $outputDirectory 'test-results.log'
$started = Get-Date
$unit = Start-Process -FilePath $unitExe -WorkingDirectory $outputDirectory -WindowStyle Hidden -PassThru
try {
    if (-not $unit.WaitForExit(15000)) {
        $unit.Kill()
        [void]$unit.WaitForExit(5000)
        throw 'Stress engine tests exceeded the 15-second outer deadline.'
    }
    $unit.Refresh()
    $unitExitCode = $unit.ExitCode
} finally { $unit.Dispose() }
if (-not (Test-Path -LiteralPath $unitLog)) { throw 'Stress engine tests produced no report.' }
if ((Get-Item -LiteralPath $unitLog).LastWriteTime -lt $started.AddSeconds(-2)) { throw 'Stress engine report is stale.' }
$unitResult = [IO.File]::ReadAllText($unitLog)
Write-Output $unitResult.Trim()
if ($unitExitCode -ne 0 -or $unitResult -notmatch '\d+ checks; 0 failures\.' -or $unitResult -match '(?m)^(FAIL|UNEXPECTED)') {
    throw "Stress engine tests failed; see $unitLog"
}
if (-not $SkipOwnership) { & (Join-Path $PSScriptRoot 'Test-StressWorkers.ps1') -SkipBuild }
if ($SkipIntegration) { return }

Add-Type -Path (Join-Path $testSources 'StressProbe.cs')
$existing = [StressProbe]::OpenFileMapping(0xF001F, $false, $MappingName)
if ($existing -ne [IntPtr]::Zero) {
    [void][StressProbe]::CloseHandle($existing)
    throw 'Production integration tests require no existing CPU Stress session; its shared mapping is already open. No existing application was changed.'
}
if ($WorkerCount -eq 0) { $WorkerCount = [Math]::Min(32, [Environment]::ProcessorCount) }
$appExe = if ($AppExecutable) { [IO.Path]::GetFullPath($AppExecutable) } else { Join-Path $root 'build\stress-integration\CPU_Stress.exe' }
$main = $null
$workers = @()
$mapping = [IntPtr]::Zero
$memory = [IntPtr]::Zero
$job = [IntPtr]::Zero
$window = [IntPtr]::Zero
$script:integrationChecks = 0
$integrationLog = Join-Path $outputDirectory ('integration-' + (Get-Date -Format 'yyyyMMdd-HHmmssfff') + '.log')
$lines = New-Object 'System.Collections.Generic.List[string]'
function Record([string]$Text) { $lines.Add($Text); Write-Host $Text }
function Assert-Stress([bool]$Condition, [string]$Description) {
    if (-not $Condition) { throw "FAIL: $Description" }
    $script:integrationChecks++
    Record "PASS $Description"
}
function Wait-Until([scriptblock]$Predicate, [int]$Milliseconds, [string]$Description) {
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (& $Predicate) { return }
        if ($deadline.ElapsedMilliseconds -ge $Milliseconds) { throw "Timed out: $Description" }
        Start-Sleep -Milliseconds 10
    } while ($true)
}
function Command-Workers([int]$Command) {
    foreach ($worker in $workers) { [StressProbe]::WriteSlot($memory, $worker.Slot, 2, $Command) }
}

try {
    $job = [StressProbe]::CreateJobObject([IntPtr]::Zero, $null)
    Assert-Stress ($job -ne [IntPtr]::Zero) 'test process job created'
    $jobLimits = New-Object byte[] $(if ([IntPtr]::Size -eq 8) { 144 } else { 112 })
    [BitConverter]::GetBytes([uint32]0x2000).CopyTo($jobLimits, 16)
    Assert-Stress ([StressProbe]::SetInformationJobObject($job, 9, $jobLimits, $jobLimits.Length)) 'test workers die when the test job closes'
    $main = Start-Process -FilePath $appExe -WorkingDirectory (Split-Path $appExe -Parent) -WindowStyle Hidden -PassThru
    Assert-Stress ([StressProbe]::AssignProcessToJobObject($job, $main.Handle)) 'test owns the main application process'
    Wait-Until {
        $script:mapping = [StressProbe]::OpenFileMapping(0xF001F, $false, $MappingName)
        $script:mapping -ne [IntPtr]::Zero
    } 5000 'main application shared mapping'
    $memory = [StressProbe]::MapViewOfFile($mapping, 0xF001F, 0, 0, [UIntPtr]::new([uint32]8192))
    Assert-Stress ($memory -ne [IntPtr]::Zero) 'mapped test application control slots'
    Wait-Until { [StressProbe]::ReadSlot($memory, 0, 0) -eq $main.Id } 5000 'main process registration'
    Wait-Until {
        $script:window = [StressProbe]::FindForm([uint32]$main.Id)
        $script:window -ne [IntPtr]::Zero
    } 5000 'test application window'
    $main.Refresh()
    Assert-Stress ($main.PriorityClass -eq [Diagnostics.ProcessPriorityClass]::AboveNormal) 'main application has AboveNormal priority'

    # Launch workers idle, so the load interval starts only after every process
    # has initialized. Commands touch only their command field, not whole slots.
    for ($slot = 1; $slot -le $WorkerCount; $slot++) {
        [StressProbe]::WriteSlot($memory, $slot, 0, -1)
        [StressProbe]::WriteSlot($memory, $slot, 1, $slot)
        [StressProbe]::WriteSlot($memory, $slot, 2, 0)
        $process = Start-Process -FilePath $appExe -ArgumentList "uigbfhs;uygbrf:$slot" -WorkingDirectory (Split-Path $appExe -Parent) -WindowStyle Hidden -PassThru
        $workers += [pscustomobject]@{ Slot = $slot; Process = $process; BeforeSample = $null; AfterSample = $null }
        Assert-Stress ([StressProbe]::AssignProcessToJobObject($job, $process.Handle)) "worker $slot belongs to the test job"
    }
    Wait-Until {
        $registered = $true
        foreach ($worker in $workers) {
            if ([StressProbe]::ReadSlot($memory, $worker.Slot, 0) -ne $worker.Process.Id) { $registered = $false }
        }
        $registered
    } 10000 'all worker registrations'
    foreach ($worker in $workers) {
        $worker.Process.Refresh()
        Assert-Stress ($worker.Process.PriorityClass -eq [Diagnostics.ProcessPriorityClass]::Normal) "worker $($worker.Slot) has Normal priority"
    }
    Command-Workers 5
    Wait-Until {
        $running = $true
        foreach ($worker in $workers) {
            if ([StressProbe]::ReadSlot($memory, $worker.Slot, 3) -ne 0xA2) { $running = $false }
        }
        $running
    } 3000 'all workers running'
    foreach ($worker in $workers) {
        $worker.BeforeSample = [StressProbe]::SampleCpu($worker.Process.Handle)
    }
    $loadTimer = [Diagnostics.Stopwatch]::StartNew()
    $maxPingMs = 0.0
    $pings = 0
    do {
        $reply = [UIntPtr]::Zero
        $pingTimer = [Diagnostics.Stopwatch]::StartNew()
        $answered = [StressProbe]::SendMessageTimeout($window, 0, [UIntPtr]::Zero, [IntPtr]::Zero, 2, $DiagnosticPingTimeoutMs, [ref]$reply)
        $pingError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        $pingTimer.Stop()
        if ($answered -eq [IntPtr]::Zero) {
            $main.Refresh()
            Record ("UI probe failed after {0:N2} ms: Windows error {1}; main exited={2}; window={3}." -f $pingTimer.Elapsed.TotalMilliseconds, $pingError, $main.HasExited, $window)
        }
        Assert-Stress ($answered -ne [IntPtr]::Zero) 'main event loop answers WM_NULL during full load'
        $maxPingMs = [Math]::Max($maxPingMs, $pingTimer.Elapsed.TotalMilliseconds)
        $pings++
        Start-Sleep -Milliseconds 100
    } while ($loadTimer.ElapsedMilliseconds -lt 3000)
    $loadTimer.Stop()
    # Capture all counters before printing diagnostics, then stop immediately.
    # Console output under full load must not extend later workers' intervals.
    foreach ($worker in $workers) { $worker.AfterSample = [StressProbe]::SampleCpu($worker.Process.Handle) }
    $stopTimer = [Diagnostics.Stopwatch]::StartNew()
    Command-Workers 7
    Wait-Until {
        $stopped = $true
        foreach ($worker in $workers) {
            if ([StressProbe]::ReadSlot($memory, $worker.Slot, 3) -ne 0xA0) { $stopped = $false }
        }
        $stopped
    } 1500 'all workers acknowledge stop'
    $stopTimer.Stop()
    $summedUtilization = 0.0
    $minimumSampleMs = [double]::PositiveInfinity
    $maximumSampleMs = 0.0
    foreach ($worker in $workers) {
        $afterSample = $worker.AfterSample
        $usedMs = $afterSample.CpuMilliseconds - $worker.BeforeSample.CpuMilliseconds
        $sampleMs = $afterSample.TimestampMilliseconds - $worker.BeforeSample.TimestampMilliseconds
        Assert-Stress ($usedMs -gt 0) "worker $($worker.Slot) consumed CPU time"
        $summedUtilization += 100 * $usedMs / $sampleMs
        $minimumSampleMs = [Math]::Min($minimumSampleMs, $sampleMs)
        $maximumSampleMs = [Math]::Max($maximumSampleMs, $sampleMs)
    }
    Record ("Measured {0} workers: {1:N1}% aggregate worker CPU capacity; per-worker sample windows {2:N0}..{3:N0} ms; max UI reply {4:N2} ms ({5} probes)." -f $WorkerCount, ($summedUtilization / $WorkerCount), $minimumSampleMs, $maximumSampleMs, $maxPingMs, $pings)
    Assert-Stress ($maxPingMs -lt 300) 'main event loop stays within the 300 ms responsiveness budget'
    Assert-Stress ($stopTimer.ElapsedMilliseconds -lt 1500) 'all workers stop within 1.5 seconds'
    Record ("All-worker stop acknowledgement: {0:N2} ms." -f $stopTimer.Elapsed.TotalMilliseconds)
    Command-Workers 17
    foreach ($worker in $workers) {
        Assert-Stress ($worker.Process.WaitForExit(3000)) "worker $($worker.Slot) exits normally"
        Assert-Stress ([StressProbe]::ReadSlot($memory, $worker.Slot, 0) -eq 0) "worker $($worker.Slot) releases its shared slot"
    }
    [void][StressProbe]::PostMessage($window, 0x10, [UIntPtr]::Zero, [IntPtr]::Zero)
    Assert-Stress ($main.WaitForExit(3000)) 'main application closes after stress'
    Record "$($script:integrationChecks) integration checks; 0 failures."
} catch {
    Record $_.Exception.Message
    throw
} finally {
    if ($memory -ne [IntPtr]::Zero) {
        foreach ($worker in $workers) { [StressProbe]::WriteSlot($memory, $worker.Slot, 2, 17) }
    }
    if ($job -ne [IntPtr]::Zero) { [void][StressProbe]::CloseHandle($job) }
    foreach ($worker in $workers) {
        try {
            if (-not $worker.Process.HasExited) { $worker.Process.Kill(); [void]$worker.Process.WaitForExit(2000) }
        } finally { $worker.Process.Dispose() }
    }
    if ($null -ne $main) {
        try {
            if (-not $main.HasExited) { $main.Kill(); [void]$main.WaitForExit(2000) }
        } finally { $main.Dispose() }
    }
    if ($memory -ne [IntPtr]::Zero) { [void][StressProbe]::UnmapViewOfFile($memory) }
    if ($mapping -ne [IntPtr]::Zero) { [void][StressProbe]::CloseHandle($mapping) }
    [IO.File]::WriteAllLines($integrationLog, $lines)
    Write-Output "Integration report: $integrationLog"
}
