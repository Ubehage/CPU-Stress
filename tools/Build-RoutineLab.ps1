param(
    [string]$NativeTools = 'C:\Users\Ubehage\Documents\WinLibs\mingw64\bin'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$sourcePath = Join-Path $root 'routine-lab'
$outputPath = Join-Path $root 'build\routine-lab'
$routineOutput = Join-Path $outputPath 'routines'
$nasm = Join-Path $NativeTools 'nasm.exe'
$gcc = Join-Path $NativeTools 'gcc.exe'
foreach ($toolPath in @($nasm, $gcc)) {
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        throw "Required native build tool not found: $toolPath"
    }
}
New-Item -ItemType Directory -Force -Path $routineOutput | Out-Null

# Native tools are invoked explicitly so compiler diagnostics are preserved.
foreach ($sample in Get-ChildItem -LiteralPath (Join-Path $sourcePath 'routines') -Filter '*.asm' -File) {
    $binaryPath = Join-Path $routineOutput ($sample.BaseName + '.bin')
    & $nasm -f bin -Wall -Werror -o $binaryPath $sample.FullName
    if ($LASTEXITCODE -ne 0) { throw "NASM failed for $($sample.Name)" }
    $bytes = [IO.File]::ReadAllBytes($binaryPath)
    $hex = ($bytes | ForEach-Object { $_.ToString('X2') }) -join ' '
    [IO.File]::WriteAllText((Join-Path $routineOutput ($sample.BaseName + '.hex')), $hex + "`r`n", [Text.Encoding]::ASCII)
}

$gateSource = Join-Path $sourcePath 'native\wow64-gate.asm'
& $nasm -f bin -Wall -Werror -o (Join-Path $outputPath 'wow64-gate.bin') $gateSource
if ($LASTEXITCODE -ne 0) { throw 'NASM failed for the WOW64 mode-switch wrapper.' }

$workerSource = Join-Path $sourcePath 'native\worker64.c'
& $gcc -std=c11 -O2 -Wall -Wextra -Werror -static -s -o (Join-Path $outputPath 'RoutineWorker64.exe') $workerSource
if ($LASTEXITCODE -ne 0) { throw 'GCC failed for RoutineWorker64.exe.' }

# VB6 source files need ANSI text and Windows line endings.
$vbEncoding = [Text.Encoding]::GetEncoding(1252)
Get-ChildItem -LiteralPath $sourcePath -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.vbp' } | ForEach-Object {
    $contents = [IO.File]::ReadAllText($_.FullName, $vbEncoding)
    $contents = [regex]::Replace($contents, '\r?\n', "`r`n")
    [IO.File]::WriteAllText($_.FullName, $contents, $vbEncoding)
}
& (Join-Path $PSScriptRoot 'Build-VB6.ps1') -Project 'routine-lab\RoutineLab.vbp' -OutputDirectory 'build\routine-lab'
Write-Output "Routine lab built in $outputPath"
