---
name: development
description: Work on the ExcelCSIToolBoxAddIn solution safely. Use when changing architecture, services, ETABS/SAP2000 integration, project references, build setup, tests, or repo organization in this VSTO/.NET Framework solution.
---

# Development

## Repo Shape

Treat the solution as layered:

- `src/ExcelCSIToolBox.Core`: pure shared models, result types, abstractions, and common logic. Do not reference ETABS/SAP2000 COM, Excel Interop, WPF, or VSTO.
- `src/ExcelCSIToolBox.Application`: use cases, workflows, and interface contracts. Do not place direct COM or Excel Interop calls here.
- `src/ExcelCSIToolBox.Infrastructure`: ETABS/SAP2000 API adapters, Excel Interop implementations, file/system integrations, and concrete service implementations.
- `src/ExcelCSIToolBoxAddIn`: VSTO add-in host, ribbon, WPF views, ViewModels, task panes, and dependency wiring.
- `src/ExcelCSIToolBox.Data`: legacy DTO/data model project kept only during migration.
- `tests/ExcelCSIToolBox.Tests`: focused unit tests.
- `tools/ExcelCSIToolBox.RefBuilder`: development-time CSI API reference/index utility.

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

## Toolbox Sections

Keep these toolbox areas conceptually separate:

- `Analysis Results`: table/result extraction through registry, router, and handlers.
- `Modelling Helper`: modelling utilities and geometry/helper workflows, currently centered on `UI/ViewModels/CsiToolboxViewModel.ModellingHelper.cs` and related helper windows.
- `Miscellaneous Data`: ETABS project/material data exports. Keep menu metadata distinct from Analysis Results even when the UI shares the same content tab surface.
- `Element Manipulation`: point/frame/shell/object-connectivity tools and selection workflows.

Do not move Modelling Helper commands into Analysis Results handlers unless the user explicitly asks for a shared action architecture. If adding a new helper action, follow the existing helper command/window pattern first, then extract services only when backend logic grows.

For Miscellaneous Data, preserve the left-tree category label and breadcrumb category. If it needs routing later, create a separate registry/router or a clearly named shared table export service rather than hiding it inside Analysis Results-specific classes.

## Build And Test

Use Visual Studio MSBuild if `msbuild` is not on PATH:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' ExcelCSIToolBox.sln /t:Build /p:Configuration=Debug /p:Platform='Any CPU' /m
```

Run tests:

```powershell
dotnet test tests\ExcelCSIToolBox.Tests\ExcelCSIToolBox.Tests.csproj --configuration Debug --no-build
```

If a change touches WPF XAML, VSTO project includes, project references, or generated resources, run the full solution build.
