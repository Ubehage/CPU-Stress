param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project 'tests\MouseTests.vbp' -OutputDirectory 'build\tests'
}
$exe = Join-Path $root 'build\tests\MouseTests.exe'
$log = Join-Path $root 'build\tests\test-results.log'
$started = Get-Date
# The harness shows only its own test form and restores the cursor afterward.
$process = Start-Process -FilePath $exe -WorkingDirectory $root -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(30000)) {
    throw "Mouse tests did not finish in 30 seconds (PID $($process.Id))."
}
if (-not (Test-Path -LiteralPath $log)) { throw 'The test program produced no result log.' }
if ((Get-Item -LiteralPath $log).LastWriteTime -lt $started.AddSeconds(-2)) {
    throw 'The test result log was not updated by this run.'
}
$result = Get-Content -LiteralPath $log -Raw
Write-Output $result.Trim()
if ($result -match '(?m)^(FAIL |UNEXPECTED:)' -or $result -notmatch '\d+ checks; 0 failures\.') {
    throw "Mouse integration tests did not pass. See $log"
}
