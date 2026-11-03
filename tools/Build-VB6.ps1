param(
    [string]$Project = 'CPU_Stress.vbp',
    [string]$OutputDirectory = 'build\app'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$projectPath = (Resolve-Path (Join-Path $root $Project)).Path
$outputPath = [IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
if (-not $outputPath.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Build output must be inside the project workspace.'
}
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$logPath = Join-Path $outputPath ('compile-' + (Get-Date -Format 'yyyyMMdd-HHmmssfff') + '.log')
$compiler = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'
$arguments = '/make "{0}" /out "{1}" /outdir "{2}"' -f $projectPath, $logPath, $outputPath
$process = Start-Process -FilePath $compiler -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru
if (-not $process.WaitForExit(60000)) {
    throw "VB6 did not finish in 60 seconds (PID $($process.Id)); inspect the compiler before retrying."
}
if (-not (Test-Path -LiteralPath $logPath)) { throw "VB6 produced no build log: $logPath" }
$result = Get-Content -LiteralPath $logPath -Raw
Write-Output $result.Trim()
if ($result -notmatch "Build of '.+' succeeded\.") { throw "VB6 build failed; see $logPath" }
