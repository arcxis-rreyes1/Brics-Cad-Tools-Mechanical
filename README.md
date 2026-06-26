# Arcxis Cad Tools - Brics (Mechanical)

Mechanical and MEP-focused BricsCAD plugin forked from [Brics-Cad-Tools](https://github.com/arcxis-rreyes1/Brics-Cad-Tools-).

## Scope

This repository keeps mechanical/MEP tooling and shared infrastructure. Framing, tendon tagging, attic vent, take-off, and title-block layout tools from the original repo were removed.

### Key commands

| Command | Description |
|---------|-------------|
| `AMP` | Open mechanical printing setup |
| `overnightprinting` | Headless mechanical batch printing (requires `PLAN TYPE` = `MECHANICAL`) |

## Requirements

- BricsCAD V26 (x64)
- .NET 8 SDK
- Visual Studio 2022+
- Microsoft Excel (for some mechanical utilities)

## Build

1. Open `Arcxis Cad Tools - Brics.sln`
2. Confirm BricsCAD DLL paths in the `.vbproj` match your install location
3. Build **Debug \| x64** or **Release \| x64**
4. Load the output DLL in BricsCAD with `NETLOAD`

## Original repo

The full multi-discipline plugin remains in the parent Brics-Cad-Tools repository.
