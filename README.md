# ExcelCSIToolBoxAddIn

ExcelCSIToolBoxAddIn is a Microsoft Excel VSTO add-in that integrates Excel workflows with CSI ETABS through the ETABS v1 API. It provides a toolbox window where engineers can connect to a running ETABS instance, read model metadata, push point data from Excel into ETABS, and pull selected ETABS point data back into Excel.

## What this repository contains

- **Excel add-in shell**: Ribbon button and add-in startup glue for launching the ETABS toolbox window.
- **WPF toolbox UI**: A tabbed window for Point/Frame operations.
- **Application use-cases**: Orchestrated operations such as attach to ETABS, close ETABS, select points by unique name, add points by Cartesian coordinates, and export selected point data.
- **Infrastructure adapters**: ETABS API service wrappers and Excel range I/O services.
- **Shared primitives**: Relay command and operation result patterns used across UI and core logic.

## Current feature status

### Implemented

- Attach to a running ETABS instance.
- Display active model name/path and current model units.
- Close the currently attached ETABS instance.
- Select ETABS points based on unique names from an Excel range.
- Add ETABS points from Excel Cartesian coordinate rows.
- Export selected ETABS point information to Excel.

### Placeholder / in progress

- Additional Point actions (set/rename/grouping variants).
- Most Frame tab actions are currently placeholders.

## Solution structure

```text
AddIn/                Excel add-in window launch + window lifetime helpers
Common/               Reusable command/result abstractions
Core/Application/     Use-case orchestration logic
Core/Tabular/         Lightweight tabular model helpers
Infrastructure/Etabs/ ETABS API connection + model operations
Infrastructure/Excel/ Excel selection/output services
UI/ViewModels/        MVVM state and commands
UI/Views/             WPF toolbox window
```

## Prerequisites

- Windows with Microsoft Excel installed.
- .NET Framework 4.8 developer tooling.
- Visual Studio with Office/SharePoint development workload (VSTO support).
- A compatible ETABS installation and API access.
- `ETABSv1.dll` available to the project (already included in this repository).

## Build and run

1. Open `ExcelCSIToolBoxAddIn.sln` in Visual Studio.
2. Restore/build the solution using `Debug|AnyCPU` or `Release|AnyCPU`.
3. Start debugging from Visual Studio to launch Excel with the add-in loaded.
4. In Excel, use the custom ribbon button to open the **ETABS Toolbox** window.

## Notes for contributors

- This project targets **.NET Framework 4.8** and uses classic `.csproj`/VSTO project style.
- UI is implemented with WPF + MVVM-style view models.
- ETABS interaction is mediated through `IEtabsConnectionService` and application use-cases to keep UI orchestration simple.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
