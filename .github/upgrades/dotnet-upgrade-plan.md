# .NET 8.0 Upgrade Plan

## Execution Steps

Execute steps below sequentially one by one in the order they are listed.

1. Validate that an .NET 8.0 SDK required for this upgrade is installed on the machine and if not, help to get it installed.
2. Ensure that the SDK version specified in global.json files is compatible with the .NET 8.0 upgrade.
3. Upgrade Arcxis Cad Tools\Arcxis Cad Tools.vbproj
5. Run unit tests to validate upgrade in the projects listed below:


## Settings

This section contains settings and data used by execution steps.

### Excluded projects

Table below contains projects that do belong to the dependency graph for selected projects and should not be included in the upgrade.

| Project name                                   | Description                 |
|:-----------------------------------------------|:---------------------------:|
|                                                |                             |

### Aggregate NuGet packages modifications across all projects

NuGet packages used across all selected projects or their dependencies that need version update in projects that reference them.

| Package Name                        | Current Version | New Version | Description                                   |
|:------------------------------------|:---------------:|:-----------:|:----------------------------------------------|
| AutoCAD.NET                         |   24.3.0        |  25.1.0     | Incompatible with .NET 8, update to compatible version |
| PDFsharp                            |   1.50.5147     |  6.2.3      | Incompatible with .NET 8, update to compatible version |
| System.Buffers                      |   4.5.1         |             | Functionality included with framework reference; remove |
| System.Numerics.Vectors             |   4.5.0         |             | Functionality included with framework reference; remove |

### Project upgrade details
This section contains details about each project upgrade and modifications that need to be done in the project.

#### Arcxis Cad Tools\Arcxis Cad Tools.vbproj modifications

Project properties changes:
  - Target framework should be changed from `.NETFramework,Version=v4.8` to `net8.0-windows`

NuGet packages changes:
  - `AutoCAD.NET` should be updated from `24.3.0` to `25.1.0` (*recommended for .NET 8*)
  - `PDFsharp` should be updated from `1.50.5147` to `6.2.3` (*recommended for .NET 8*)
  - `System.Buffers` should be removed (*functionality included in framework reference*)
  - `System.Numerics.Vectors` should be removed (*functionality included in framework reference*)

Feature upgrades:
  - Convert project file to SDK-style format.

Other changes:
  - Ensure WPF-specific Windows target framework is used (`net8.0-windows`) and UI references updated accordingly.
