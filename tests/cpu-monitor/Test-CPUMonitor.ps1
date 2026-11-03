$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$sourceFolder = Join-Path $root 'build\monitor-test-source'
$outputFolder = Join-Path $root 'build\monitor-tests'
[void][IO.Directory]::CreateDirectory($sourceFolder)
[void][IO.Directory]::CreateDirectory($outputFolder)
$encoding = [Text.Encoding]::GetEncoding(1252)
$projectLines = [IO.File]::ReadAllLines((Join-Path $root 'tests\MouseTests.vbp'), $encoding)
$newLines = foreach ($line in $projectLines) {
    if ($line -match '^(Form|UserControl|Module|Class)=(.*)$') {
        $kind = $Matches[1]
        $value = $Matches[2]
        $parts = $value.Split(';', 2)
        $path = $parts[$parts.Length - 1].Trim()
        $fullPath = [IO.Path]::GetFullPath((Join-Path (Join-Path $root 'tests') $path))
        $name = [IO.Path]::GetFileName($fullPath)
        Copy-Item -LiteralPath $fullPath -Destination (Join-Path $sourceFolder $name) -Force
        foreach ($extension in @('.ctx', '.frx')) {
            $sidecar = [IO.Path]::ChangeExtension($fullPath, $extension)
            if (Test-Path -LiteralPath $sidecar) { Copy-Item -LiteralPath $sidecar -Destination $sourceFolder -Force }
        }
        if ($parts.Length -eq 2) { $kind + '=' + $parts[0] + '; ' + $name } else { $kind + '=' + $name }
    } elseif ($line -match '^ExeName32=') { 'ExeName32="MonitorTests.exe"' }
    else { $line }
}
$newLines += 'Module=modMonitorTests; modMonitorTests.bas'
[IO.File]::WriteAllLines((Join-Path $sourceFolder 'MonitorTests.vbp'), $newLines, $encoding)
$mouseMain = Join-Path $sourceFolder 'modMouseTests.bas'
$text = [IO.File]::ReadAllText($mouseMain, $encoding).Replace('Public Sub Main()', 'Public Sub UnusedMouseTestMain()')
[IO.File]::WriteAllText($mouseMain, $text, $encoding)
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'modMonitorTests.bas') -Destination $sourceFolder -Force
$controlPath = Join-Path $sourceFolder 'CPUView.ctl'
$control = [IO.File]::ReadAllText($controlPath, $encoding)
if ($control -match '(?i)Pdh(OpenQuery|AddCounter|CollectQueryData)') { throw 'CPU monitor still contains blocking PDH setup.' }
# Accessors exist only in this generated test copy, never in production code.
$control += @'

Public Sub MonitorTestSample()
  GetCPULoads
End Sub
Public Function MonitorTestHasData() As Boolean
  MonitorTestHasData = gHasData
End Function
Public Function MonitorTestBaseline() As Boolean
  MonitorTestBaseline = m_HaveBaseline
End Function
Public Function MonitorTestUnavailable() As Boolean
  MonitorTestUnavailable = m_LoadUnavailable
End Function
Public Function MonitorTestLoad(ByVal Index As Long) As Double
  MonitorTestLoad = CPULoad(Index)
End Function
Public Sub MonitorTestReset()
  ShiftAutoUpdate
End Sub
Public Sub MonitorTestDisableQuery()
  m_QueryResolved = True
  m_QueryAddress = 0
End Sub
Public Sub MonitorTestRestoreQuery()
  m_QueryResolved = False
End Sub
'@
[IO.File]::WriteAllText($controlPath, [regex]::Replace($control, '\r?\n', "`r`n"), $encoding)
$modulePath = Join-Path $sourceFolder 'modMonitorTests.bas'
[IO.File]::WriteAllText($modulePath, [regex]::Replace([IO.File]::ReadAllText($modulePath, $encoding), '\r?\n', "`r`n"), $encoding)
& (Join-Path $root 'tools\Build-VB6.ps1') -Project 'build\monitor-test-source\MonitorTests.vbp' -OutputDirectory 'build\monitor-tests'
$process = Start-Process -FilePath (Join-Path $outputFolder 'MonitorTests.exe') -WorkingDirectory $outputFolder -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(10000)) { $process.Kill(); throw 'CPU monitor test timed out.' }
$report = [IO.File]::ReadAllText((Join-Path $outputFolder 'monitor-results.log'))
Write-Output $report.Trim()
if ($process.ExitCode -ne 0 -or $report -notmatch '\d+ checks; 0 failures\.' -or $report -match '(?m)^(FAIL |UNEXPECTED:)') {
    throw 'CPU monitor query-path checks failed.'
}
