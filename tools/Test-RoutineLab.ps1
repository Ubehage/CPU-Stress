param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot 'Build-RoutineLab.ps1')
}
$labDirectory = Join-Path $root 'build\routine-lab'
$controller = Join-Path $labDirectory 'RoutineLab.exe'
$nativeWorker = Join-Path $labDirectory 'RoutineWorker64.exe'
$samples = Join-Path $labDirectory 'routines'
$runDirectory = Join-Path $labDirectory ('test-runs\' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($runDirectory)
$script:checks = 0
$script:skips = 0

function Assert-Lab([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message (reports: $runDirectory)" }
    $script:checks++
}

function Get-PEMachine([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $header = [BitConverter]::ToInt32($bytes, 0x3c)
    Assert-Lab ($bytes[0] -eq 0x4d -and $bytes[1] -eq 0x5a) "$Path has an MZ header"
    Assert-Lab ([BitConverter]::ToUInt32($bytes, $header) -eq 0x4550) "$Path has a PE header"
    return [BitConverter]::ToUInt16($bytes, $header + 4)
}

function New-HexFixture([string]$Name, [string]$Text) {
    $path = Join-Path $runDirectory ($Name + '.hex')
    [IO.File]::WriteAllText($path, $Text, [Text.Encoding]::ASCII)
    return $path
}

function Invoke-LabCase {
    param(
        [string]$Name, [string]$Mode, [string]$HexPath,
        [int]$Iterations = 7, [int]$TimeoutMs = 3000,
        [string]$Requirement = 'base'
    )
    $report = Join-Path $runDirectory ($Name + '.txt')
    foreach ($path in @($HexPath, $report)) {
        if ($path.Contains('"')) { throw 'Test paths must not contain quotes.' }
    }
    $arguments = '/run {0} "{1}" {2} {3} "{4}" {5}' -f $Mode, $HexPath, $Iterations, $TimeoutMs, $report, $Requirement
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $process = Start-Process -FilePath $controller -ArgumentList $arguments -WorkingDirectory $labDirectory -WindowStyle Hidden -PassThru
    try {
        # The controller has its own worker deadline. This second limit also
        # catches bugs in the controller and keeps automation from hanging.
        if (-not $process.WaitForExit(15000)) {
            $process.Kill()
            [void]$process.WaitForExit(5000)
            throw "Controller exceeded the 15-second outer deadline: $Name"
        }
        $process.Refresh()
        $exitCode = $process.ExitCode
    } finally {
        $timer.Stop()
        $process.Dispose()
    }
    Assert-Lab (Test-Path -LiteralPath $report) "$Name wrote a report"
    $fields = @{}
    foreach ($line in [IO.File]::ReadAllLines($report)) {
        $separator = $line.IndexOf('=')
        if ($separator -gt 0) { $fields[$line.Substring(0, $separator)] = $line.Substring($separator + 1) }
    }
    Assert-Lab ($fields.ContainsKey('status')) "$Name reported status"
    Assert-Lab (($exitCode -eq 0) -eq ($fields.status -eq 'ok')) "$Name exit code agrees with status"
    Write-Host ("{0}: {1} ({2:N0} ms wall time)" -f $Name, $fields.status, $timer.Elapsed.TotalMilliseconds)
    return [pscustomobject]@{ Fields = $fields; ExitCode = $exitCode; WallMs = $timer.Elapsed.TotalMilliseconds }
}

function Assert-OK($Result, [string]$Mode) {
    Assert-Lab ($Result.Fields.status -eq 'ok') "$Mode routine completed: $($Result.Fields.error)"
    $architecture = if ($Mode -eq 'wow64') { 'wow64-x64' } else { $Mode }
    Assert-Lab ($Result.Fields.architecture -eq $architecture) "$Mode reports the execution architecture"
    Assert-Lab ($Result.Fields.scratch_hex -match '^[0-9A-Fa-f]{32}$') "$Mode reports sixteen scratch bytes"
    $elapsed = 0.0
    Assert-Lab ([double]::TryParse($Result.Fields.elapsed_ms, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$elapsed) -and $elapsed -ge 0) "$Mode reports elapsed milliseconds"
    if ($Mode -eq 'wow64') {
        Assert-Lab ($Result.Fields.code_segment -match '^(0x)?0*33$') 'WOW64 proof executed with code selector 0x33'
    }
}

function Assert-FiniteScratch($Result, [string]$Name) {
    $hex = $Result.Fields.scratch_hex
    $bytes = New-Object byte[] 16
    for ($index = 0; $index -lt 16; $index++) { $bytes[$index] = [Convert]::ToByte($hex.Substring(2 * $index, 2), 16) }
    $nonzero = $false
    for ($lane = 0; $lane -lt 4; $lane++) {
        $value = [BitConverter]::ToSingle($bytes, 4 * $lane)
        Assert-Lab (-not [single]::IsNaN($value) -and -not [single]::IsInfinity($value)) "$Name lane $lane is finite"
        if ($value -ne 0) { $nonzero = $true }
    }
    Assert-Lab $nonzero "$Name writes a nonzero arithmetic result"
}

function Get-ReferenceScratch([int]$Iterations, [bool]$Fused) {
    # Independent scalar model of the sample recurrence. Each product of
    # two binary32 inputs fits exactly in binary64 for these bounded inputs.
    # SSE rounds the product before adding; FMA rounds only the final sum.
    [single]$value = 0.5
    [single]$factor = 1.0 - [Math]::Pow(2.0, -24)
    [single]$bias = [Math]::Pow(2.0, -23)
    for ($iteration = 0; $iteration -lt $Iterations; $iteration++) {
        [double]$product = [double]$value * [double]$factor
        if (-not $Fused) { $product = [double][single]$product }
        $value = [single]($product + [double]$bias)
    }
    Assert-Lab ($value -gt 0.5) "Scalar reference evolves beyond its initial value after $Iterations iterations"
    $lane = ([BitConverter]::GetBytes([single]$value) | ForEach-Object { $_.ToString('X2') }) -join ''
    return ($lane * 4)
}

Assert-Lab ((Get-PEMachine $controller) -eq 0x014c) 'Controller and WOW64 worker are PE32 x86'
Assert-Lab ((Get-PEMachine $nativeWorker) -eq 0x8664) 'Comparison worker is PE32+ AMD64'

$return42 = Join-Path $samples 'return42-x86.hex'
$proof64 = Join-Path $samples 'proof64-x64.hex'
$result = Invoke-LabCase 'return42-x86' 'x86' $return42
Assert-OK $result 'x86'
Assert-Lab ($result.Fields.return_hex -eq '0000002A') 'x86 return value is 42'

foreach ($mode in @('wow64', 'x64')) {
    $result = Invoke-LabCase "proof64-$mode" $mode $proof64
    Assert-OK $result $mode
    Assert-Lab ($result.Fields.return_hex -eq '1122334555667788') "$mode retains the entire 64-bit result"
    Assert-Lab ($result.Fields.scratch_hex.StartsWith('8877665545332211')) "$mode stores the 64-bit result in little-endian scratch"
}

$size32 = New-HexFixture 'scratch-size-x86' '8B 44 24 10 C2 10 00'
$size64 = New-HexFixture 'scratch-size-x64' '4C 89 C8 C3'
foreach ($mode in @('x86', 'wow64', 'x64')) {
    $fixture = if ($mode -eq 'x86') { $size32 } else { $size64 }
    $result = Invoke-LabCase "scratch-size-$mode" $mode $fixture
    Assert-OK $result $mode
    $expectedSize = if ($mode -eq 'wow64') { 4080 } else { 4096 }
    Assert-Lab ([Convert]::ToUInt64($result.Fields.return_hex, 16) -eq $expectedSize) "$mode receives the usable scratch size"
}

$alignment64 = New-HexFixture 'stack-alignment-x64' '48 89 E0 83 E0 0F C3'
foreach ($mode in @('wow64', 'x64')) {
    $result = Invoke-LabCase "stack-alignment-$mode" $mode $alignment64
    Assert-OK $result $mode
    Assert-Lab ([Convert]::ToUInt64($result.Fields.return_hex, 16) -eq 8) "$mode stack is 8 modulo 16 at routine entry"
}

foreach ($requirement in @('sse2', 'fma')) {
    $modes = if ($requirement -eq 'sse2') { @('x86', 'wow64', 'x64') } else { @('wow64', 'x64') }
    foreach ($mode in $modes) {
        $suffix = if ($mode -eq 'x86') { 'x86' } else { 'x64' }
        $sample = Join-Path $samples ("$requirement-$suffix.hex")
        $firstScratch = $null
        foreach ($iterations in @(1, 31)) {
            $name = "$requirement-$mode-$iterations"
            $result = Invoke-LabCase $name $mode $sample -Iterations $iterations -Requirement $requirement
            if ($requirement -eq 'fma' -and $result.Fields.status -eq 'unsupported') {
                $script:skips++
                Write-Host "SKIP: FMA is unavailable for $mode."
                break
            }
            Assert-OK $result $mode
            Assert-FiniteScratch $result $name
            Assert-Lab ([Convert]::ToUInt64($result.Fields.return_hex, 16) -eq $iterations) "$name completes the requested iterations"
            $expectedScratch = Get-ReferenceScratch $iterations ($requirement -eq 'fma')
            Assert-Lab ($result.Fields.scratch_hex -eq $expectedScratch) "$name matches the independently rounded scalar recurrence"
            if ($iterations -eq 1) {
                $firstScratch = $result.Fields.scratch_hex
            } else {
                Assert-Lab ($result.Fields.scratch_hex -ne $firstScratch) "$name advances beyond the one-iteration output"
            }
        }
    }
}

$invalid = New-HexFixture 'invalid' 'B8 ZZ C3'
$fault = New-HexFixture 'invalid-instruction' '0F 0B'
$loop = New-HexFixture 'infinite-loop' 'EB FE'
foreach ($mode in @('x86', 'wow64', 'x64')) {
    $good = if ($mode -eq 'x86') { $return42 } else { $proof64 }
    $result = Invoke-LabCase "invalid-hex-$mode" $mode $invalid
    Assert-Lab ($result.Fields.status -eq 'error') "$mode rejects invalid hex"
    $result = Invoke-LabCase "zero-count-$mode" $mode $good -Iterations 0
    Assert-Lab ($result.Fields.status -eq 'error') "$mode rejects zero iterations"
    $result = Invoke-LabCase "ud2-$mode" $mode $fault
    Assert-Lab ($result.Fields.status -in @('exception', 'crash')) "$mode contains the invalid-instruction fault"
    $result = Invoke-LabCase "timeout-$mode" $mode $loop -TimeoutMs 250
    Assert-Lab ($result.Fields.status -eq 'timeout') "$mode terminates an endless loop"
    Assert-Lab ($result.WallMs -lt 10000) "$mode deadline finishes promptly"
    $result = Invoke-LabCase "recovery-$mode" $mode $good
    Assert-OK $result $mode
    $expectedReturn = if ($mode -eq 'x86') { '0000002A' } else { '1122334555667788' }
    Assert-Lab ($result.Fields.return_hex -eq $expectedReturn) "$mode runs successfully after fault and timeout"
}

Write-Output "$($script:checks) checks passed; $($script:skips) unsupported-feature skips. Reports: $runDirectory"
