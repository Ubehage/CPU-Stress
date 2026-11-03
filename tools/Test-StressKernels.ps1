param(
    [string]$NativeTools = 'C:\Users\Ubehage\Documents\WinLibs\mingw64\bin'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$labDirectory = Join-Path $root 'build\routine-lab'
$controller = Join-Path $labDirectory 'RoutineLab.exe'
if (-not (Test-Path -LiteralPath $controller -PathType Leaf)) {
    & (Join-Path $PSScriptRoot 'Build-RoutineLab.ps1') -NativeTools $NativeTools
}
& (Join-Path $PSScriptRoot 'Generate-StressCode.ps1') -Check -NativeTools $NativeTools
$runDirectory = Join-Path $root ('build\stress-code\runs\' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($runDirectory)
$script:checks = 0
$script:skips = 0
$results = @{}

function Assert-Kernel([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message (reports: $runDirectory)" }
    $script:checks++
}

function ConvertTo-HexFile([string]$BinaryPath, [string]$Name) {
    $hexPath = Join-Path $runDirectory ($Name + '.hex')
    $hex = ([IO.File]::ReadAllBytes($BinaryPath) | ForEach-Object { $_.ToString('X2') }) -join ' '
    [IO.File]::WriteAllText($hexPath, $hex + "`r`n", [Text.Encoding]::ASCII)
    return $hexPath
}

function Invoke-Kernel([string]$Name, [string]$Mode, [string]$HexPath, [int]$Iterations, [string]$Requirement) {
    $reportPath = Join-Path $runDirectory ($Name + '.txt')
    $arguments = '/run {0} "{1}" {2} 3000 "{3}" {4}' -f $Mode, $HexPath, $Iterations, $reportPath, $Requirement
    $process = Start-Process -FilePath $controller -ArgumentList $arguments -WorkingDirectory $labDirectory -WindowStyle Hidden -PassThru
    try {
        if (-not $process.WaitForExit(10000)) {
            $process.Kill()
            [void]$process.WaitForExit(3000)
            throw "Routine controller exceeded its 10-second outer deadline: $Name"
        }
        $process.Refresh()
        $exitCode = $process.ExitCode
    } finally {
        $process.Dispose()
    }
    Assert-Kernel (Test-Path -LiteralPath $reportPath -PathType Leaf) "$Name produced a report"
    $fields = @{}
    foreach ($line in [IO.File]::ReadAllLines($reportPath)) {
        $separator = $line.IndexOf('=')
        if ($separator -gt 0) { $fields[$line.Substring(0, $separator)] = $line.Substring($separator + 1) }
    }
    Assert-Kernel ($fields.ContainsKey('status')) "$Name reported status"
    Assert-Kernel (($exitCode -eq 0) -eq ($fields.status -eq 'ok')) "$Name exit code agrees with its report"
    Write-Host "$Name`: $($fields.status)"
    return $fields
}

function Assert-FiniteScratch($Fields, [string]$Name) {
    Assert-Kernel ($Fields.scratch_hex -match '^[0-9A-Fa-f]{32}$') "$Name reports sixteen scratch bytes"
    $bytes = New-Object byte[] 16
    for ($index = 0; $index -lt 16; $index++) { $bytes[$index] = [Convert]::ToByte($Fields.scratch_hex.Substring(2 * $index, 2), 16) }
    for ($lane = 0; $lane -lt 4; $lane++) {
        $value = [BitConverter]::ToSingle($bytes, 4 * $lane)
        Assert-Kernel (-not [single]::IsNaN($value) -and -not [single]::IsInfinity($value)) "$Name lane $lane is finite"
        Assert-Kernel ($value -gt 0.01 -and $value -lt 2.0) "$Name lane $lane stays in its bounded normal range"
    }
}

function Get-ReferenceScratch([int]$Iterations, [bool]$Fused) {
    # Products of two binary32 operands fit exactly in binary64. For these
    # bounded values the bias addition also fits before rounding to binary32.
    # Separate multiply/add explicitly rounds the product first; FMA does not.
    [single]$factorDown = 1.0 - [Math]::Pow(2.0, -24)
    [single]$factorUp = 1.0 + [Math]::Pow(2.0, -23)
    [single]$biasUp = [Math]::Pow(2.0, -20)
    [single]$biasDown = -[Math]::Pow(2.0, -21)
    $effectiveIterations = (($Iterations - 1) % 256) + 1
    $hex = ''
    for ($lane = 0; $lane -lt 4; $lane++) {
        [single]$value = 0.125 + $lane / 64.0
        for ($iteration = 0; $iteration -lt $effectiveIterations; $iteration++) {
            for ($step = 0; $step -lt 4; $step++) {
                $factor = if (($step % 2) -eq 0) { $factorDown } else { $factorUp }
                $bias = if (($step % 2) -eq 0) { $biasUp } else { $biasDown }
                [double]$product = [double]$value * [double]$factor
                if (-not $Fused) { $product = [double][single]$product }
                $value = [single]($product + [double]$bias)
            }
        }
        $hex += ([BitConverter]::GetBytes($value) | ForEach-Object { $_.ToString('X2') }) -join ''
    }
    return $hex
}

$cases = @(
    @{ Kernel = 'fma64'; Requirement = 'fma'; Modes = @('wow64', 'x64') },
    @{ Kernel = 'avx64'; Requirement = 'avx'; Modes = @('wow64', 'x64') },
    @{ Kernel = 'sse64'; Requirement = 'sse2'; Modes = @('wow64', 'x64') },
    @{ Kernel = 'sse32'; Requirement = 'sse2'; Modes = @('x86') },
    @{ Kernel = 'legacy32'; Requirement = 'base'; Modes = @('x86') }
)
foreach ($case in $cases) {
    $hexPath = ConvertTo-HexFile (Join-Path $root ('build\stress-code\' + $case.Kernel + '.bin')) $case.Kernel
    foreach ($mode in $case.Modes) {
        $firstScratch = $null
        foreach ($iterations in @(1, 31, 256, 257)) {
            $name = "$($case.Kernel)-$mode-$iterations"
            $fields = Invoke-Kernel $name $mode $hexPath $iterations $case.Requirement
            if ($fields.status -eq 'unsupported' -and $case.Requirement -in @('avx', 'fma')) {
                Assert-Kernel ($iterations -eq 1) "$name consistently rejects an unsupported feature"
                $script:skips++
                break
            }
            Assert-Kernel ($fields.status -eq 'ok') "$name completed: $($fields.error)"
            $architecture = if ($mode -eq 'wow64') { 'wow64-x64' } else { $mode }
            Assert-Kernel ($fields.architecture -eq $architecture) "$name reports its execution architecture"
            Assert-Kernel ([Convert]::ToUInt64($fields.return_hex, 16) -eq $iterations) "$name returns the requested count"
            if ($mode -eq 'wow64') { Assert-Kernel ($fields.code_segment -match '^(0x)?0*33$') "$name entered the 64-bit code segment" }
            Assert-FiniteScratch $fields $name
            if ($case.Kernel -ne 'legacy32') {
                $reference = Get-ReferenceScratch $iterations ($case.Kernel -eq 'fma64')
                Assert-Kernel ($fields.scratch_hex -eq $reference) "$name matches the independently rounded scalar reference"
            }
            if ($iterations -eq 1) { $firstScratch = $fields.scratch_hex }
            if ($iterations -eq 31) { Assert-Kernel ($fields.scratch_hex -ne $firstScratch) "$name evolves beyond the one-iteration output" }
            if ($iterations -eq 257) { Assert-Kernel ($fields.scratch_hex -eq $firstScratch) "$name restarts seeds after 256 iterations" }
            $results[$name] = $fields
        }
    }
    if ($case.Modes.Count -eq 2) {
        foreach ($iterations in @(1, 31, 256, 257)) {
            $wowName = "$($case.Kernel)-wow64-$iterations"
            $nativeName = "$($case.Kernel)-x64-$iterations"
            Assert-Kernel ($results.ContainsKey($wowName) -eq $results.ContainsKey($nativeName)) "$($case.Kernel) capability agrees between WOW64 and native execution"
            if ($results.ContainsKey($wowName)) {
                Assert-Kernel ($results[$wowName].scratch_hex -eq $results[$nativeName].scratch_hex) "$($case.Kernel) iteration $iterations matches between WOW64 and native execution"
            }
        }
    }
}

# Call the exact legacy kernel eight times while keeping an x87 sentinel live.
# This detects both stack leaks and accidental popping of the caller's value.
$legacyBinary = (Join-Path $root 'build\stress-code\legacy32.bin').Replace('\', '/')
$sentinelAsm = @'
BITS 32
push ebp
mov ebp, esp
push ebx
push esi
push edi
mov ebx, [ebp+12]
fnclex
fnstsw ax
and eax, 0x3800
mov esi, eax
fld1
mov edi, 8
.again:
push dword [ebp+20]
push dword [ebp+16]
push ebx
push dword [ebp+8]
call kernel
dec edi
jnz .again
fstp dword [ebx]
fnstsw ax
movzx eax, ax
mov edx, eax
and eax, 0x3800
xor eax, esi
and edx, 0x41
or eax, edx
pop edi
pop esi
pop ebx
pop ebp
ret 16
kernel:
'@
$sentinelAsm += "`r`nincbin `"$legacyBinary`"`r`n"
$sentinelSource = Join-Path $runDirectory 'legacy-stack-sentinel.asm'
$sentinelBinary = Join-Path $runDirectory 'legacy-stack-sentinel.bin'
[IO.File]::WriteAllText($sentinelSource, $sentinelAsm, [Text.Encoding]::ASCII)
& (Join-Path $NativeTools 'nasm.exe') -f bin -Wall -Werror -o $sentinelBinary $sentinelSource
if ($LASTEXITCODE -ne 0) { throw 'Cannot assemble the legacy x87 stack sentinel.' }
$sentinelHex = ConvertTo-HexFile $sentinelBinary 'legacy-stack-sentinel'
$sentinel = Invoke-Kernel 'legacy-stack-sentinel' 'x86' $sentinelHex 257 'base'
Assert-Kernel ($sentinel.status -eq 'ok') 'Legacy stack sentinel executes successfully'
Assert-Kernel ($sentinel.return_hex -eq '00000000') 'Legacy kernel preserves x87 TOP and avoids stack/invalid-operation faults across eight calls'
Assert-Kernel ($sentinel.scratch_hex.StartsWith('0000803F')) 'Legacy kernel preserves the caller x87 sentinel value'

Write-Output "$($script:checks) checks passed; $($script:skips) unsupported-feature skips. Reports: $runDirectory"
