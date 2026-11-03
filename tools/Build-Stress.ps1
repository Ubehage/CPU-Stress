param([switch]$TestIsolation)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
& (Join-Path $PSScriptRoot 'Generate-StressCode.ps1') -Check
$encoding = [Text.Encoding]::GetEncoding(1252)
foreach ($name in @('CPU_Stress.vbp', 'modStress.bas', 'modStressCode.bas', 'modStressWorkers.bas', 'modMain.bas', 'modSystem.bas', 'modCPU.bas', 'modSharedMemory.bas', 'frmMain.frm', 'frmStress.frm', 'CPUView.ctl')) {
    $path = Join-Path $root $name
    $text = [IO.File]::ReadAllText($path, $encoding)
    [IO.File]::WriteAllText($path, [regex]::Replace($text, '\r?\n', "`r`n"), $encoding)
}
# Compile an isolated source snapshot so an open IDE project is untouched.
$snapshotRelative = if ($TestIsolation) { 'build\stress-test-source' } else { 'build\stress-source' }
$snapshot = Join-Path $root $snapshotRelative
[void][IO.Directory]::CreateDirectory($snapshot)
foreach ($file in Get-ChildItem -LiteralPath $root -File | Where-Object { $_.Extension.ToLowerInvariant() -in '.bas', '.cls', '.ctl', '.ctx', '.frm', '.frx', '.vbp', '.res' }) {
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $snapshot $file.Name) -Force
}
$snapshotProject = Join-Path $snapshot 'CPU_Stress.vbp'
$projectText = [IO.File]::ReadAllText($snapshotProject, $encoding)
$projectText = [regex]::Replace($projectText, '(?m)^Reference=.*OLE Automation\r?$', 'Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation')
[IO.File]::WriteAllText($snapshotProject, $projectText, $encoding)
if ($TestIsolation) {
    # Only the test snapshot has a separate control mapping/settings namespace.
    # It exercises the same compiled logic without touching an open IDE/app.
    $sharedPath = Join-Path $snapshot 'modSharedMemory.bas'
    $sharedText = [IO.File]::ReadAllText($sharedPath, $encoding).Replace('Local\UbeCPUStress', 'Local\UbeCPUStressIntegration')
    [IO.File]::WriteAllText($sharedPath, $sharedText, $encoding)
    $mainPath = Join-Path $snapshot 'modMain.bas'
    $mainText = [IO.File]::ReadAllText($mainPath, $encoding).Replace('"UbesCPUStress"', '"UbesCPUStressIntegration"')
    [IO.File]::WriteAllText($mainPath, $mainText, $encoding)
    $formPath = Join-Path $snapshot 'frmMain.frm'
    $formText = [IO.File]::ReadAllText($formPath, $encoding).Replace('  SaveUISettings', "  ' Test snapshot does not save desktop preferences.")
    [IO.File]::WriteAllText($formPath, $formText, $encoding)
}
$outputRelative = if ($TestIsolation) { 'build\stress-integration' } else { 'build\app' }
& (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project ($snapshotRelative + '\CPU_Stress.vbp') -OutputDirectory $outputRelative
