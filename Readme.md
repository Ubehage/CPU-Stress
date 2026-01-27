# CPU-Stress Tool (Work-In-Progress / Experimental)

Status: WIP · Experimental

A minimal, focused CPU stress & control tool with a compact UI that lets the user control every logical execution unit (kernel) individually.
The architecture and design will mirror my MemEater project where it makes sense: Minimal runtime dependencies, explicit user control, predictable behavior, and telemetry for safety.

This repository is in early development.
The UI is fully functional, but all the core functions for stressing the cpu is not yet implemented.

---


## Current status
- Visual view of the processor cores.
- Live update of the cpu's workload.
- Customizable delay between updates.
- No stress-generation code implemented yet.

## Vision
- A compact, responsive CPU-stress application with a minimalistic GUI that exposes:
  - Per-core and per-kernel control (start/stop stress test).
  - User-configurable profiles (safe, test, destructive) and a visible kill-switch.
  - No built-in safeguards. My vision is to make this tool a cpu-killer, and give the user full control and responsibility.
  - No external dependencies - fully portable executable.
- Follow the same architectural design principles used in MemEater: multi processes, communicating via shared memory.

## Safety & usage guidance
- This project will include code that intentionally overloads CPUs. Do NOT run stress features on hardware you do not own or in production environments.
- For now the repository is safe. When stress features are added, they will be explicitly labeled and documented.
- Recommended test environments (when stress code exists): isolated VMs, test rigs, or otherwise non-critical hardware with active thermal monitoring.

## Screenshot
![The UI so far](./screenshot.png)
## License
- MIT License. All code is free to use and modify.

## Why source-only now?
- Why not?