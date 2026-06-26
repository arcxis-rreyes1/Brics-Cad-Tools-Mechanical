# Arcxis Cad Tools - Brics - Mechanical

Mechanical/MEP BricsCAD plugin trimmed from the full Brics-Cad-Tools repo.

## Scope

This repository contains mechanical printing and shared submission utilities only:

| Command | Description |
|---------|-------------|
| `AMP` | Interactive mechanical printing (MechAutoPage layouts → PDF + CSV) |
| `overnightprinting` | Headless mechanical batch printing (`PLAN TYPE` must be `MECHANICAL`) |
| `B2C` | Save a copy of the current DWG to Egnyte for AutoCAD conversion |
| `xsr` | Convert xref paths to absolute and resolve missing xrefs |
| `RedoPaths` | Refresh Arcxis support file search paths |

## Requirements

- BricsCAD V26 (x64)
- .NET 8 SDK
- Visual Studio 2022+

## Build

1. Open `Arcxis Cad Tools - Brics - Mechanical.sln`
2. Confirm BricsCAD DLL paths in the `.vbproj` match your install location
3. Build **Debug \| x64** or **Release \| x64**
4. Load the output DLL in BricsCAD with `NETLOAD`

## Original repo

The full multi-discipline plugin remains in the parent Brics-Cad-Tools repository.
