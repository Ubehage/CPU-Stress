# Stress loop and native routine execution

Audit date: 2026-11-03. Scope: `modStress.bas`, `modCPU.bas`, and `frmStress.frm`. The findings below describe the original production implementation, before any subsequent fixes. No original stress routine was executed for this audit.

The current mechanism can run native machine code, but several byte sequences do something different from their comments. Correctness, repeatable inputs, and bounded execution should come before comparing their load on the CPU. The user's clarified objective is real 64-bit instructions inside the 32-bit process. An experimental WOW64 transition can address that objective; a separate native x64 worker remains the conventional supported alternative.

## How the existing mechanism works

`SplitAssemblyBytes` converts a hexadecimal string into a VB6 byte array. `AllowAssemblyExecution` changes its pages to executable/readable/writable. `CallWindowProcA` then calls the array's address as a window procedure. This executes ordinary x86 machine instructions in the existing process; there is no bytecode interpreter or injection into another process.

The routine receives four 32-bit arguments. After `push ebp / mov ebp,esp`, `[ebp+16]` contains the third argument, `wParam`, which supplies the iteration count. `ret 16` removes those four arguments. The actual instruction sequences preserve the nonvolatile general registers they modify. Their comments are less reliable: the SSE2 comment includes a nonexistent `push ebx`, while the legacy epilogue comment omits the actual `pop ebx`. The x86 convention requires preserving EBX, ESI, EDI and EBP when used. [Microsoft calling conventions](https://learn.microsoft.com/en-us/cpp/cpp/argument-passing-and-naming-conventions?view=msvc-170).

`StressLoop` caches its byte array and address, then requests 1,000,000,000 iterations synchronously. `frmStress.RunStress` checks stop, hibernation, and parent-process state only after this call returns. Responsiveness therefore depends on the entire batch duration. A zero count would also enter the body and wrap the decrementing counter; a future routine contract needs explicit zero handling or strict positive-count validation.

## What the bytes actually do

The exact source strings were extracted and decoded with `ndisasm -b 32`. Reproduce with:

```powershell
rtk proxy powershell -NoProfile -File build\stress-audit\disassemble.ps1
```

The corresponding `.bin` and `.disasm.txt` files are under `build/stress-audit`. Offsets below are hexadecimal, relative to the beginning of each routine.

| Selected routine | Finding | Consequence |
| --- | --- | --- |
| `ccAVX_YMM` | `75 C9` at `0x66` branches to `0x31`, whereas the first multiply is at `0x2D`. | The first multiply executes only on the first iteration. |
| `ccAVX_YMM` | For example, `C5 FC 59 D1` is `vmulps ymm2,ymm0,ymm1`, not the commented `vmulps ymm2,ymm2,ymm1`. The other destinations likewise read `ymm0`; the adds have the same issue. | The intended independent accumulators are absent. Most destination results are overwritten without feeding their next iteration. |
| `ccAVX` | This routine uses legacy SSE/SSE2 encodings, despite its AVX name. `75 D1` at `0x51` targets `0x24`, inside `movaps xmm7,xmm1` at `0x23`. | At that target, bytes `28 F9` decode as `sub cl,bh`, followed by the arithmetic at `0x26`. This is unintended execution, though it need not crash. |
| `ccSSE2` | `75 F1` at `0x20` correctly targets the first multiply at `0x13`. All arithmetic updates `xmm0`. | One dependent accumulator limits instruction overlap; fixing the branch is unnecessary here. |
| `ccLegacy` | `DE F1` decodes as `fdivrp st1`, rather than the commented `fdivp`. Two `fld1` instructions are followed by three divide-and-pop instructions and another pop. | With an initially empty x87 stack, the second divide already lacks its second operand. The loop underflows the stack and does not maintain balanced state. |

For the unchanged layouts, the intended AVX branch displacement would be `75 C5`, and the SSE branch would be `75 D3`. These are diagnostic observations, not a complete repair: operand encoding and initialization also need correction. Keep assembly with labels as the source of truth and generate bytes with an assembler.

All three SIMD routines initialize with scalar **double** conversion and then calculate with packed **single** arithmetic. Double `1.0` has bits `3FF0000000000000`; interpreting its two halves as singles gives `0.0` and `1.875`, not two copies of `1.0`. Legacy `cvtsi2sd` leaves the upper 64 bits of XMM unchanged. The VEX form copies those bits from its first source, which here is the previous `xmm1`, and clears the upper YMM half. Thus some lanes are uninitialized by the routine and others are zero. SSE multiplication repeatedly grows the known nonzero lane toward overflow. [Intel CVTSI2SD instruction reference](https://www.intel.com/content/dam/www/public/us/en/documents/manuals/64-ia-32-architectures-software-developer-vol-2a-manual.pdf).

The AVX operand decoding follows the VEX source-register field, independently of the destination register. Packed single multiplication operates lane by lane. [Intel MULPS/VMULPS instruction reference](https://www.intel.com/content/dam/www/public/us/en/documents/manuals/64-ia-32-architectures-software-developer-vol-2b-manual.pdf). The x87 divide-and-pop behavior explains the stack accounting above. [Intel FDIVR/FDIVRP instruction reference](https://www.intel.la/content/dam/www/public/us/en/documents/manuals/itanium-architecture-vol-4-manual.pdf), IA-32 instruction-reference section.

## Execution support and memory lifecycle

`GetCPUCapabilities` checks AVX hardware support before issuing `XGETBV`, but never checks OSXSAVE. AVX hardware support alone does not establish that the OS enabled `XGETBV`. Require CPUID leaf 1 ECX OSXSAVE bit 27 and AVX bit 28, then check `(XGETBV(0) & 6) == 6`. Checking XSAVE bit 26 explicitly is also reasonable; Intel states that OSXSAVE already implies hardware XSAVE support. FMA needs its own feature check, and AVX2 requires leaf 7 EBX bit 5 in addition to usable AVX. [Intel feature-detection guidance](https://cdrdv2-public.intel.com/821613/355308-Software-Optimization-Manual-048-Changes-Doc.pdf).

`AllowAssemblyExecution` ignores `VirtualProtect` failure and changes pages belonging to a VB-managed allocation. Multiple allocations may share those pages; local CPUID arrays also disappear without restoring their pages' original protection. Microsoft advises avoiding page-protection changes on individual heap allocations. Use dedicated virtual allocation, copy while writable, change to executable/readable, check every API result, and release the allocation after execution. [Microsoft VirtualProtect](https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-virtualprotect).

Flush the process instruction cache after installing executable bytes, as required for generated or modified code. The current helper omits this step. [Microsoft FlushInstructionCache](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-flushinstructioncache).

The current parser also rejects a single-byte routine and does not handle repeated whitespace or line breaks robustly. A work program should report malformed tokens before allocating or invoking anything.

## A routine work program and genuine x64 execution

Use a controller with an explicit architecture choice, byte input, bounded iteration count, timeout, and result/timing display. Execute experimental routines in disposable worker processes so a bad return, invalid instruction, or endless loop does not take down the controller. Process isolation is fault containment, not a security sandbox for untrusted code.

The x86 worker can retain the existing four-argument callback contract. For the requested in-process experiment, that callback can enter a transition stub, execute a small x64 routine, store its complete 64-bit result, switch back, and return to VB6. A disposable **32-bit** worker can contain this entire experiment: that still demonstrates x64 execution inside a 32-bit process. A result exceeding 32 bits and a recorded code selector provide useful evidence of the transition.

ReWolf's original WOW64Ext implementation demonstrates far-return transitions using code selectors `0x33` to enter 64-bit code and `0x23` to return to 32-bit code. These Windows implementation details are not a portable application ABI. The example establishes that in-process execution is possible; it does not establish compatibility with every Windows build, mitigation configuration, or non-x64 host. [Original transition macros](https://github.com/rwfpl/rewolf-wow64ext/blob/master/src/internal.h).

Keep the first experiment limited to local arithmetic leaf routines, with no Windows calls, loader interaction, or exceptions crossing the mode boundary. Define and preserve register and stack state explicitly, zero-extend 32-bit addresses, and restore the x86 return context exactly. The original library's call bridge also manages stack alignment and FS/SS restoration, and identifies R12 as carrying WOW64 thread state. These are concrete reasons to test the whole round trip instead of treating a far transfer alone as a finished bridge. [Original X64Call implementation](https://github.com/rwfpl/rewolf-wow64ext/blob/master/src/wow64ext.cpp).

The resulting [Routine Lab](routine-lab/README.md) passed 310 local end-to-end checks on 2026-11-03. Its PE32 worker entered code selector `0033`, performed 64-bit arithmetic, returned `1122334555667788`, and stored the same value in scratch memory. SSE2/FMA outputs at one and 31 iterations matched independent scalar reference calculations and evolved with the count; deliberate faults and endless loops were isolated and subsequent runs succeeded. These measurements establish the requested direct execution on this machine, without asserting broad Windows compatibility or changing the production stress loop.

For a conventional alternative, Microsoft supports communication between 32-bit and 64-bit processes; a 32-bit application cannot load an ordinary 64-bit DLL. Launching an x64 EXE with an explicit IPC or file protocol avoids reliance on those transition details, but does not satisfy the user's specific same-process experiment by itself. [Microsoft process interoperability](https://learn.microsoft.com/en-us/windows/win32/winprog64/process-interoperability).

Windows x64 integer arguments arrive in RCX, RDX, R8, and R9. The caller provides shadow space; stack alignment and nonvolatile-register rules differ from x86. In particular, XMM6-XMM15 have nonvolatile low 128-bit portions. Initially prefer leaf routines using volatile registers, with no external calls or stack manipulation; arbitrary non-leaf generated code additionally needs correct unwind support. [Microsoft x64 calling convention](https://learn.microsoft.com/en-us/cpp/build/x64-calling-convention?view=msvc-170).

AVX vector width and process bitness are separate choices: the 256-bit VMULPS form is valid in both 32-bit and 64-bit modes when CPU and OS support it. x64 enables 64-bit general registers, addressing and the x64 ABI; it is not required merely to use YMM arithmetic. [Intel VMULPS mode-support table](https://www.intel.com/content/dam/www/public/us/en/documents/manuals/64-ia-32-architectures-software-developer-vol-2b-manual.pdf).

## Moving toward a stronger stress load

First compare correct, deterministic kernels: bounded finite input values, several independent accumulators, deliberate instruction mix, and a checkable output. Use short batches with cancellation checks between them; measure and tune batch duration rather than assuming one billion iterations is appropriate on every CPU. Keep a separate latency-oriented dependent-chain test if that is useful, but do not confuse it with a throughput test.

Measure operations per second separately from sustained package power, temperature, clock rate, and throttling. A fully scheduled thread can show 100% utilization while executing a relatively weak or dependency-limited workload. Wider arithmetic is a candidate to measure, not proof of maximum heat: Intel documents that AVX workloads can reduce operating frequency because of power characteristics. [Intel AVX power and frequency guidance](https://www.intel.com/content/www/us/en/support/articles/000090746/processors.html).

The existing affinity path uses a 32-bit mask, handles indices only through 31, and ignores affinity-call failure. Before expanding worker counts, make placement explicit and verify success; broader topology and processor-group support should be designed into the native worker separately.

This audit establishes instruction and control-flow defects, not a thermal benchmark. It does not claim that any proposed kernel is the hottest workload on the user's CPU. The work program provides the place to establish correctness first and make that comparison empirically.

## Production x64 stress engine (2026-11-03)

The production application now embeds the transition bridge and replacement kernels in `modStressCode.bas`. Every instruction is represented by space-separated hexadecimal bytes alongside its assembly comment. NASM sources live in `stress-kernels`; `tools/Generate-StressCode.ps1 -Check` verifies both the exact bytes and their comments. The compiled application needs no external assembly files, bytecode files, or native helper executable.

On AMD64 Windows, each existing 32-bit worker enters x64 mode locally. The preferred FMA kernel uses twelve independent YMM accumulators and four updates per loop iteration: 768 single-precision floating-point operations. Distinct lane/register seeds are refreshed every 256 iterations, keeping values finite and changing rather than drifting into infinities or a rounded fixed point. AVX multiply/add and SSE2 alternatives are included, with corrected x86 SSE2 and balanced x87 fallbacks. CPU and OS state are checked before choosing AVX/FMA.

`modStress.bas` owns dedicated executable allocations and scratch memory. It starts at 4096 iterations, then adapts batch size toward 15 ms. The eight-million-iteration cap and calibration bound ordinary batches; they are not a hard real-time deadline under suspension or extreme scheduling delays. Workers check commands and parent lifetime between batches. Commands and status now update separate aligned shared-memory fields so a status acknowledgement cannot overwrite a newer Stop request.

All selected processors can remain busy. Worker processes use Normal priority; the separate main window uses AboveNormal priority and performs no stress arithmetic. Its higher scheduling priority lets input preempt computation without leaving a processor idle. The owner holds workers in a kill-on-close job, including during suspended startup, so closing or losing the main process ends its workers. [Windows scheduling priorities](https://learn.microsoft.com/en-us/windows/win32/procthread/scheduling-priorities).

Integration testing also found a separate UI stall: the first PDH counter setup took about 1020 ms, with 660 ms spent adding the first counter. This reproduced with only two workers, so scheduling priority alone could not solve it. `CPUView` now dynamically resolves `NtQuerySystemInformation` and obtains one per-processor timing snapshot, then computes busy time from successive idle/kernel/user counters. Failed or incompatible snapshots display `N/A` and discard stale measurements. Monitoring no longer initializes PDH counters on the UI thread. [Microsoft processor performance information](https://learn.microsoft.com/en-us/windows/win32/api/winternl/nf-winternl-ntquerysysteminformation).

Build with `tools/Build-Stress.ps1`; the executable is `build/app/CPU_Stress.exe`. The script compiles an isolated source snapshot to avoid interfering with an open VB6 project. `tools/Test-Stress.ps1` tests the engine, worker ownership, and a separately named integration build; `tools/Test-StressKernels.ps1` compares all emitted kernels through the routine lab. The integration build uses a separate shared-memory name and does not save desktop preferences.

The existing processor topology/affinity implementation supports logical processor indices 0 through 31, including all 32 logical processors on the tested Ryzen 9 3950X. Affinity failures now report failure instead of silently running an unbound worker. Larger processor-group support remains a separate change. The direct WOW64 transition retains its experimental Windows-specific compatibility constraints.

Final validation passed 84 engine checks, 19 worker-lifetime checks, 189 full-application checks, 566 kernel/reference checks, 162 existing mouse/control checks, and 36 CPU-monitor checks. During the short 32-worker run, workers consumed 95.6% of aggregate logical-processor capacity; maximum main-window message latency was 12.65 ms and all-worker Stop acknowledgement was 14.80 ms. Monitoring query tests measured approximately 0.15 ms for the initial snapshot. See `tests/stress/README.md` and `build/stress-tests/integration-20261103-025532686.log` for scope and measurements. Sustained package power, temperature, and comparisons against other kernel mixes have not been measured; these results establish a working heavy FMA load and responsive control, not a universal maximum-heat claim.
