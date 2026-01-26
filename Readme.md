# CPU-Stress Tool (Work-In-Progress / Experimental)

Status: WIP · Experimental

A minimal, focused CPU stress & control tool with a compact UI that lets the user control every logical execution unit (kernel) individually.
The architecture and design will mirror my MemEater project where it makes sense: Minimal runtime dependencies, explicit user control, predictable behavior, and telemetry for safety.

This repository is an early-stage snapshot:
Only CPU discovery/enumeration is implemented so far.
The intent of this public repository is to share design intent, invite targeted design feedback, and track progress as the project evolves.

---

## What this repo contains (right now)
- UserControl with a visual overview of all the cpu-cores and how busy they are.
- A simple form, showing this usercontrol.

## Current status — implemented
- Processor enumeration: physical core count, per-core logical kernels reported.
- UserControl with a visual overview of all the cpu-cores and how busy they are.
- A simple form, showing this usercontrol.
- No stress-generation code implemented yet.

## Why publish at this stage?
- Early feedback on the approach and cross-Windows compatibility can prevent architectural rework later.

## Vision
- A compact, responsive CPU-stress application with a minimalistic GUI that exposes:
  - Per-core and per-kernel control (start/stop stress test).
  - User-configurable profiles (safe, test, destructive) and a visible kill-switch.
  - No built-in safeguards. My vision is to make this tool a cpu-killer, and give the user full control and responsibility.
  - No external dependencies - fully portable executable.
- Follow the same architectural design principles used in MemEater: multi processes, communicating via shared memory.

## Safety & usage guidance
- This project will include code that intentionally overloads CPUs. Do NOT run stress features on hardware you do not own or in production environments.
- For now the repository is safe: it only enumerates CPU topology. When stress features are added, they will be explicitly labeled and documented.
- Recommended test environments (when stress code exists): isolated VMs, test rigs, or otherwise non-critical hardware with active thermal monitoring.

## License
- MIT License. All code is free to use and modify.

## Why source-only now?
- Why not?