---
name: development
description: Work on the ExcelCSIToolBoxAddIn solution safely. Use when changing architecture, services, ETABS/SAP2000 integration, project references, build setup, tests, or repo organization in this VSTO/.NET Framework solution.
---

# Development

## Repo Shape

Treat the solution as layered:

- `ExcelCSIToolBox.Core`: pure shared models, result types, abstractions, and common logic. Do not reference ETABS/SAP2000 COM, Excel Interop, WPF, or VSTO.
- `ExcelCSIToolBox.Application`: use cases, workflows, and interface contracts. Do not place direct COM or Excel Interop calls here.
- `ExcelCSIToolBox.Infrastructure`: ETABS/SAP2000 API adapters, Excel Interop implementations, file/system integrations, and concrete service implementations.
- `ExcelCSIToolBoxAddIn`: VSTO add-in host, ribbon, WPF views, ViewModels, task panes, and dependency wiring.
- `ExcelCSIToolBox.Data`: DTOs, data models, and table/data-frame structures.
- `ExcelCSIToolBox.Tests`: focused unit tests.

## Development Rules

- Keep ViewModels thin: UI state, binding properties, commands, validation messages, and delegation only.
- Put reusable application logic in Application use cases or contracts.
- Put ETABS/SAP2000 COM calls only in Infrastructure.
- Put Excel Interop output logic in Infrastructure unless an existing add-in-specific workflow already owns it.
- Avoid circular references. Direction should be AddIn -> Infrastructure/Application/Core, Infrastructure -> Application/Core, Application -> Core, Core -> none.
- Use existing `OperationResult` patterns where the codebase already uses them.
- Respect current attach/close/lock/unit workflows when changing ETABS toolbox behavior.
- For old-style `.csproj` files with explicit includes, update the project file when adding source files.

## Analysis Results Pattern

For Analysis Results:

- Core owns `AnalysisResultItem`, `AnalysisResultGroup`, and `EtabsTableResult`.
- Application owns interfaces such as `IEtabsAnalysisResultRouter`, `IEtabsAnalysisResultHandler`, `IEtabsDatabaseTableService`, `IEtabsUnitService`, and `IExcelExportService`.
- Infrastructure owns `EtabsAnalysisResultRegistry`, router implementation, handlers, ETABS table service, unit service, and Excel export service.
- AddIn owns command binding and dependency wiring.

Expected flow:

```text
Button click
-> ViewModel command
-> IEtabsAnalysisResultRouter.ExecuteAsync(item)
-> matching Handler.ExecuteAsync(item)
-> IEtabsUnitService.SetPresentUnitsFromMainWindow()
-> IEtabsDatabaseTableService.GetTableAsync(item.EtabsTableName)
-> IExcelExportService.ExportTable(result)
```

Tree selection should only navigate; extraction should happen only from the result button command.

## Build And Test

Use Visual Studio MSBuild if `msbuild` is not on PATH:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' ExcelCSIToolBoxAddIn.sln /t:Build /p:Configuration=Debug /p:Platform='Any CPU' /m
```

Run tests:

```powershell
dotnet test ExcelCSIToolBox.Tests\ExcelCSIToolBox.Tests.csproj --configuration Debug --no-build
```

If a change touches WPF XAML, VSTO project includes, project references, or generated resources, run the full solution build.

