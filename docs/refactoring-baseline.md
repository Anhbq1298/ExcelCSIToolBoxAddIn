# Refactoring Baseline

Phase: 0 - Baseline and Safety Check

Date: 2026-07-12

## Baseline Status

- Solution build: passed.
- Unit tests: passed, 16/16.
- Production code changes in this phase: none.
- Baseline command used for build:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" ExcelCSIToolBoxAddIn.sln /m /p:Configuration=Debug /v:minimal
```

- Baseline command used for tests:

```powershell
dotnet test ExcelCSIToolBox.Tests\ExcelCSIToolBox.Tests.csproj --configuration Debug --no-build
```

## Current Projects

| Project | Target | Main responsibility | Key dependencies |
| --- | --- | --- | --- |
| `ExcelCSIToolBoxAddIn` | .NET Framework 4.8 VSTO | Excel add-in host, ribbon, WPF/WinForms UI, ViewModels, composition root | Core, Data, Application, Infrastructure, AI, VSTO, Excel interop |
| `ExcelCSIToolBox.Core` | `net48` | Shared abstractions, result types, commands, geometry, common models | Currently links some files from Data |
| `ExcelCSIToolBox.Data` | `net48` | DTOs, data models, table schemas, mapper models | Core |
| `ExcelCSIToolBox.Application` | `net48` | Use cases and application-facing service contracts | Core, Data |
| `ExcelCSIToolBox.Infrastructure` | `net48` | ETABS/SAP2000 adapters, Excel interop services, table extraction, concrete service implementations | Application, Core, Data, ETABSv1, SAP2000v1, Excel interop |
| `ExcelCSIToolBox.AI` | `net48` | Experimental AI/MCP planning/orchestration layer | Core, Data, Application |
| `ExcelCSIToolBox.Tests` | `net48` | Unit tests for mappers, operation results, geometry, and use cases | Core, Data, Application |
| `ExcelCSIToolBox.RefBuilder` | `net48` | Reference/scaffold generation utility | Standalone utility project |

## Project Ownership Notes

- `ExcelCSIToolBox.Core.csproj` links and compiles files from `ExcelCSIToolBox.Data`:
  - `..\ExcelCSIToolBox.Data\CSISapModel\**\*.cs`
  - `..\ExcelCSIToolBox.Data\DTOs\CSI\*.cs`
  - `..\ExcelCSIToolBox.Data\Models\*.cs`
- This matches the Phase 7 risk: file ownership and project compilation boundaries are unclear.
- `ExcelCSIToolBox.Data` also references Core, which makes the ownership model harder to reason about.

## Main Workflows

### ETABS and SAP2000 Connection Services

- ETABS main service: `ExcelCSIToolBox.Infrastructure/Etabs/EtabsConnectionService.cs`.
- SAP2000 main service: `ExcelCSIToolBox.Infrastructure/Sap2000/Sap2000ConnectionService.cs`.
- Shared UI/application contract: `ExcelCSIToolBox.Core/Abstractions/CSI/ICSISapModelConnectionService.cs`.
- Both connection services currently combine session management with many point, frame, shell, loading, result, table, and model-state operations.

### Analysis Export Workflows

- Navigation and command routing are in `UI/ViewModels/CsiToolboxViewModel.Export.cs`.
- Export option popup orchestration is in `UI/ViewModels/OutputTableExportWorkflow.cs`.
- Popup execution and Excel write calls are in `UI/ViewModels/OutputTableExportOptionsViewModel.cs`.
- ETABS analysis result registry and handlers are in `ExcelCSIToolBox.Infrastructure/Services/Etabs/AnalysisResults`.
- General display table extraction is handled by `EtabsConnectionService.GetDisplayTable(...)`.

### Selection Filtering Workflows

- Selection DTO contract: `CsiSelectedObjectDto` in `ExcelCSIToolBox.Core/Abstractions/CSI/ICsiReadOnlySelectionService.cs`.
- ETABS active selection logic is currently inside `EtabsConnectionService`, including selected point/frame/area/pier matching sets.
- SAP2000 has a separate selection path in `Sap2000ConnectionService`.
- Some selected-object filtering currently lives in UI (`CsiToolboxViewModel.Export.cs`) for object connectivity export.
- There is duplicated matching behavior for unique names, labels, object-column aliases, and selected row filtering.

### Excel Read/Write Workflows

- Excel output abstraction: `ExcelCSIToolBox.Core/Abstractions/Excel/IExcelOutputService.cs`.
- Concrete Excel writer: `ExcelCSIToolBox.Infrastructure/Excel/ExcelOutputService.cs`.
- Table export popup writes through `IExcelOutputService.WriteValuesToActiveCell(...)`.
- Some ViewModels retain or activate Excel interop `Range` objects during anchor selection and export.

### Modelling Helper Workflows

- Main partial ViewModel: `UI/ViewModels/CsiToolboxViewModel.ModellingHelper.cs`.
- Offset/path helper logic: `UI/ViewModels/CsiToolboxViewModel.OffsetFromSetOfLines.cs`.
- Helper action router: `AddIn/Modules/ModellingHelpers`.
- Modelling helper code is currently coupled to the main toolbox ViewModel and active ETABS/SAP2000 connection service.

### Shell Uniform Load Set Workflow

- UI form: `UI/Forms/ShellUniformLoadSetForm.cs`.
- Export current definitions: `UI/ViewModels/CsiToolboxViewModel.Loadings.cs`.
- ETABS table implementation: `ExcelCSIToolBox.Infrastructure/CSISapModel/ShellUniformLoadSetService/EtabsShellUniformLoadSetTableService.cs`.
- Schema resolver: `ExcelCSIToolBox.Infrastructure/CSISapModel/ShellUniformLoadSetService/ShellUniformLoadSetTableSchemaResolver.cs`.
- SAP2000 service returns "ETABS only" failures for Shell Uniform Load Sets.

## Existing Tests

Test project: `ExcelCSIToolBox.Tests`.

Current test files:

- `FrameDataFrameMapperTests.cs`
- `GetFrameSectionsUseCaseTests.cs`
- `OffsetPolylineServiceTests.cs`
- `OperationResultTests.cs`

Baseline result:

- Passed: 16
- Failed: 0
- Skipped: 0
- Duration: about 0.9 seconds

## Build Requirements

- Windows.
- Visual Studio 2022 with Office/SharePoint development workload.
- .NET Framework 4.8 Developer Pack.
- VSTO build targets and runtime.
- Microsoft Excel desktop for add-in runtime verification.
- ETABS/SAP2000 installations for connected workflow verification.
- `ETABSv1.dll` and `SAP2000v1.dll` available through the repo `lib`/build references.
- Visual Studio MSBuild is the reliable build path for the VSTO project. Plain `dotnet build` may fail on machines without VSTO targets in the .NET SDK path.

## Build and Test Results

### Build

Command:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" ExcelCSIToolBoxAddIn.sln /m /p:Configuration=Debug /v:minimal
```

Result:

- Passed.
- No build warnings were reported in the captured output.

### Tests

Command:

```powershell
dotnet test ExcelCSIToolBox.Tests\ExcelCSIToolBox.Tests.csproj --configuration Debug --no-build
```

Result:

- Passed.
- Total: 16.

## Major Technical Risks

1. `EtabsConnectionService.cs` and `Sap2000ConnectionService.cs` are large capability aggregates, mixing connection lifecycle, selection, model modification, database table extraction, and analysis results.
2. `CsiToolboxViewModel` is split into partials but remains highly coupled; major feature commands and state still share one ViewModel.
3. Selected-object identity resolution is duplicated and inconsistent across ViewModels and Infrastructure.
4. `CsiToolboxViewModel.Export.cs` currently constructs Infrastructure ETABS table services directly for connectivity export.
5. `EtabsDatabaseTableService.GetTableAsync(...)` uses `Task.Run` while holding ETABS `cSapModel`, which is a COM safety risk.
6. `ShellUniformLoadSetForm.cs` uses `Task.Run` around connection-service operations.
7. Excel interop `Range` objects are used in ViewModel/popup workflows and may become stale if workbooks close.
8. Empty display-table results can currently be converted into a `"No records found"` worksheet row in some paths.
9. Project ownership is unclear because Core compiles linked source files from Data.
10. Very large generated/reference files exist, especially `CsiMethodCatalog.generated.cs`; refactors should avoid touching generated files unless the generator is updated.

## Manual Verification Required

The following cannot be fully verified by current unit tests:

- Excel add-in load/unload through VSTO.
- Ribbon command behavior inside Excel.
- Excel workbook, worksheet, and selected-cell behavior.
- ETABS attach, active model detection, model lock/unlock, present unit changes, and database table extraction.
- SAP2000 attach and model operations.
- ETABS/SAP2000 selected-object filtering against live models.
- ETABS analysis result availability, especially when a model has not been analyzed.
- Shell Uniform Load Set read/apply/verification against live ETABS tables.
- Installer behavior and certificate/trust prompts.

## Recommended Phase 1 Starting Point

Start with a pure, testable selected-object identity model and table-field alias resolver before moving more logic out of ViewModels. The first implementation should avoid broad rewrites and should not change UI behavior.
