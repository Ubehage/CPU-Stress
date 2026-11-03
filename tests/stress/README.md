# Production stress tests

Run from the project root:

```powershell
rtk proxy powershell -NoProfile -File tools\Test-Stress.ps1
```

This builds the two small VB6 test programs and a production integration snapshot. The snapshot changes only its shared-memory name, settings namespace, and persistence of UI settings, so it can coexist with an ordinary CPU Stress session without altering that session. Tests still refuse to reuse an already-open integration mapping. The full application test uses up to 32 logical processors for a three-second measurement interval, then stops and exits its own workers.

Use `-SkipBuild` with existing binaries, `-SkipIntegration` for engine/ownership tests alone, or `-WorkerCount 2` for a smaller application test. `-DiagnosticPingTimeoutMs 3000` allows collecting a slow UI reply for diagnosis; it does **not** relax the final 300 ms responsiveness requirement. `-SkipOwnership` exists for focused diagnostic runs and omits the launcher tests.

`StressTests.vbp` compiles the actual `modStress.bas` and `modStressCode.bas`. It verifies rejected counts, full return values, code selector `0033` when WOW64 is selected, finite evolving scratch data, independent scalar reference calculations at 1, 2, 31 and 257 iterations, the 256-iteration seed reset, bounded adaptive batches, and clean shutdown/reinitialization. It tests the kernel selected on the current machine. The separate `tools/Test-StressKernels.ps1` exercises all supported assembled kernels, including fallbacks.

`WorkerTests.vbp` compiles the actual `modStressWorkers.bas` with a minimal shared-slot fixture. It verifies PID/command publication, Normal worker priority, job shutdown, relaunch, failed-launch cleanup, and termination of a child after abnormal owner death. Its children only sleep; this test creates no stress load.

The full application test starts an isolated main application and its workers. It checks AboveNormal main priority and Normal worker priority, sends bounded `WM_NULL` probes to the main form while every worker runs, measures worker CPU time, and checks stop acknowledgement and slot cleanup, including the last slot. These are programmatic responsiveness tests; they do not move the pointer, activate windows, or click controls. The worker launcher's ownership semantics are covered separately by `WorkerTests`.

The harness uses kill-on-close jobs for its own processes and outer deadlines for the compiled test programs. Reports remain under `build/stress-tests`. CPU-time utilization is recorded rather than assumed: concurrent system work and scheduler activity can reduce the workers' share. A passing three-second test establishes execution, control responsiveness, and cleanup; it is not a sustained thermal benchmark.

On 2026-11-03, the Ryzen 9 3950X run passed 84 engine assertions, 18 launcher assertions plus the abnormal-owner-exit check, and 189 application integration assertions. All 32 workers selected the production execution path and consumed CPU time. Across aligned 3048 ms sampling windows, the workers used 95.6% of aggregate logical-processor capacity; the slowest main-window response was 12.65 ms, and every worker acknowledged Stop within 14.80 ms. The engine harness selected `WOW64 x64 FMA` and verified code selector `0033`. See `build/stress-tests/integration-20261103-025532686.log` for the application measurements.
