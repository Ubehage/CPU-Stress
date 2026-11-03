param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$sources = Join-Path $root 'tests\stress'
$output = Join-Path $root 'build\stress-tests'
if (-not $SkipBuild) {
    foreach ($name in @('WorkerTests.vbp', 'modWorkerTests.bas')) {
        $path = Join-Path $sources $name
        $text = [IO.File]::ReadAllText($path)
        [IO.File]::WriteAllText($path, [regex]::Replace($text, '\r?\n', "`r`n"), [Text.Encoding]::GetEncoding(1252))
    }
    & (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project 'tests\stress\WorkerTests.vbp' -OutputDirectory 'build\stress-tests'
}
$exe = Join-Path $output 'WorkerTests.exe'
$log = Join-Path $output 'worker-results.log'
$started = Get-Date
$test = Start-Process -FilePath $exe -WorkingDirectory $output -WindowStyle Hidden -PassThru
try {
    if (-not $test.WaitForExit(15000)) {
        $test.Kill()
        [void]$test.WaitForExit(2000)
        throw 'Worker ownership tests exceeded their outer deadline.'
    }
    $test.Refresh()
    if ($test.ExitCode -ne 0) { throw "Worker ownership tests failed; see $log" }
} finally { $test.Dispose() }
if (-not (Test-Path -LiteralPath $log) -or (Get-Item -LiteralPath $log).LastWriteTime -lt $started.AddSeconds(-2)) {
    throw 'Worker ownership report is missing or stale.'
}
$result = [IO.File]::ReadAllText($log)
Write-Output $result.Trim()
if ($result -notmatch '\d+ checks; 0 failures\.' -or $result -match '(?m)^(FAIL|UNEXPECTED)') { throw 'Worker ownership assertions failed.' }

# Kill only this test's owner process, without calling its shutdown function.
# The real production launcher must ensure Windows then kills its idle child.
$pidFile = Join-Path $output 'owner-child.pid'
if (Test-Path -LiteralPath $pidFile) { Remove-Item -LiteralPath $pidFile }
$owner = Start-Process -FilePath $exe -ArgumentList '/owner' -WorkingDirectory $output -WindowStyle Hidden -PassThru
$child = $null
try {
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    $childId = 0
    while ($childId -le 0) {
        if (Test-Path -LiteralPath $pidFile) {
            try { $childId = [int][IO.File]::ReadAllText($pidFile).Trim() } catch { $childId = 0 }
            if ($childId -gt 0) { break }
        }
        if ($owner.HasExited) { throw 'Ownership-test parent exited before publishing its child.' }
        if ($deadline.ElapsedMilliseconds -gt 5000) { throw 'Ownership-test parent did not publish its child.' }
        Start-Sleep -Milliseconds 10
    }
    $child = [Diagnostics.Process]::GetProcessById($childId)
    if ($child.HasExited) { throw 'Ownership-test child exited prematurely.' }
    $owner.Kill()
    if (-not $owner.WaitForExit(2000)) { throw 'Ownership-test parent did not terminate.' }
    if (-not $child.WaitForExit(2000)) { throw 'Production worker job did not terminate its child after owner death.' }
    Write-Output 'PASS abnormal owner death terminates its child through the production worker job.'
    Add-Content -LiteralPath $log -Value 'PASS abnormal owner death terminates its child through the production worker job.'
} finally {
    if (-not $owner.HasExited) { $owner.Kill(); [void]$owner.WaitForExit(2000) }
    $owner.Dispose()
    if ($null -ne $child) {
        if (-not $child.HasExited) { $child.Kill(); [void]$child.WaitForExit(2000) }
        $child.Dispose()
    }
}
