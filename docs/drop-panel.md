# Drop Panel Modelling Helper

## Repository assessment

Drop Panel is implemented as a native feature of the existing Excel VSTO add-in. The repository remains on .NET Framework 4.8 and keeps its existing project formats, C# language settings, signing configuration, and ClickOnce/VSTO deployment flow.

The solution separates UI composition, application geometry, and CSI integration:

- `ExcelCSIToolBoxAddIn` is the VSTO/WPF host.
- `ExcelCSIToolBox.Application` owns plain DTOs, assignment signatures, and batch geometry processing.
- `ExcelCSIToolBox.Infrastructure` owns all ETABS COM/API calls.
- `ExcelCSIToolBox.Tests` owns the pure geometry tests.

The referenced ETABS interop assembly is `lib/ETABSv1.dll`, assembly/file version `2.7.0.0`. Its SHA-256 at implementation time was `C0016DF99002078C86E037DD231AD4C29431ABBC5AF32F9BFF725314A1BFF5E1`. The implementation was compiled against that exact local reference; it does not substitute or generate CSI API signatures.

The closest existing Modelling Helpers reference is **Quick Create Pile Cap and Pile**:

- `UI/Modules/ModellingHelpers/QuickCreatePileCapWindow.xaml`
- `UI/Modules/ModellingHelpers/QuickCreatePileCapViewModel.cs`
- `Infrastructure/CSI/Etabs/Modelling/EtabsPileCapCreationService.cs`

It provided the main reference for the WPF layout, preview drawing, MVVM command flow, and ETABS modelling separation. Drop Panel uses the shared `ModelessWpfWindowService` instead of the older pile-cap dialog service because ETABS selection must remain interactive.

## Reused infrastructure

The feature reuses:

- `RelayCommand` and `ViewModelBase` for MVVM.
- `ICSISapModelConnectionService` and the existing ETABS attach flow.
- `CurrentThreadCsiApiDispatcher` for the controlled COM context.
- `ModelessWpfWindowService` for Excel HWND ownership, one-window activation, and add-in shutdown cleanup.
- `CsiOperationLogger` and `AddInDiagnostics` for mutation and UI diagnostics.
- `Properties.Settings` for user-scoped persistence.
- `EtabsToolboxTheme.xaml` for WPF resources.
- `CsiTableFieldAliasResolver` and the existing database-table conventions.
- The existing Excel VSTO application instance through `Globals.ExcelCSIToolBoxAddin.Application`.

No standalone process, new WPF application loop, duplicate command framework, connection manager, logger, or settings system was introduced.

## Placement and namespaces

The VSTO-facing files are in the existing Modelling Helpers folder:

```text
src/ExcelCSIToolBoxAddIn/UI/Modules/ModellingHelpers/
|-- DropPanelWindow.xaml
|-- DropPanelWindow.xaml.cs
|-- DropPanelViewModel.cs
|-- DropPanelPreviewCanvas.cs
|-- DropPanelSettingsStore.cs
`-- DropPanelExcelLogExporter.cs
```

They follow the repository's established UI namespaces:

- `ExcelCSIToolBoxAddIn.UI.Views`
- `ExcelCSIToolBoxAddIn.UI.ViewModels`

The cross-layer implementation follows the solution architecture:

```text
src/ExcelCSIToolBox.Application/Modelling/DropPanels/
src/ExcelCSIToolBox.Application/Interfaces/Etabs/IDropPanelEtabsService.cs
src/ExcelCSIToolBox.Infrastructure/CSI/Etabs/Modelling/DropPanels/EtabsDropPanelService.cs
tests/ExcelCSIToolBox.Tests/Application/Modelling/DropPanels/DropPanelGeometryProcessorTests.cs
```

The corresponding namespaces are:

- `ExcelCSIToolBox.Application.Modelling.DropPanels`
- `ExcelCSIToolBox.Application.Interfaces.Etabs`
- `ExcelCSIToolBox.Infrastructure.CSI.Etabs.Modelling.DropPanels`
- `ExcelCSIToolBox.Tests.Application.Modelling.DropPanels`

## Ribbon and lifecycle

The tool is available at:

```text
Excel -> CSI Toolbox -> Modelling Helpers -> Drop Panel
```

The Ribbon click handler only delegates to `WindowManager.ShowDropPanelWindow`. The Helpers section in the ETABS task pane also exposes the same command. `WindowManager` maintains one active instance and activates that window on subsequent requests. The shared modeless service uses `Window.Show()`, attaches the WPF owner to the Excel window handle, does not set `Topmost`, and closes registered windows during add-in shutdown.

The feature borrows the shared ETABS connection and never releases the shared ETABS or Excel root COM objects.

## Workflow and safety boundary

The workflow is deliberately split into read, pure geometry, and mutation phases:

```text
ETABS COM context
  -> selected frames, column heads, slabs, openings, and assignments as DTOs
  -> pure NetTopologySuite batch processing on a background task
  -> validated preview plan
  -> ETABS COM context for backup, replacement, restoration, and verification
```

Only pure DTO geometry is sent to `Task.Run`. No ETABS COM object is used by a worker thread.

Preview does not modify ETABS. Apply is enabled only for a complete valid plan. Immediately before mutation, Apply rechecks:

- the active model path;
- ETABS present units;
- the selected drop property;
- every source area property;
- every source area polygon coordinate.

Any change requires a new Preview. Cancellation is available during reading and geometry processing, and is disabled as soon as ETABS modification starts.

By default, Apply saves the current model and copies it beside the original as:

```text
<model>.drop-panel-backup-YYYYMMDD-HHMMSS.edb
```

The backup option is visible and defaults to enabled. If it is disabled, the confirmation explicitly warns that automatic rollback is unavailable. After deletion begins, an exception triggers automatic file-backup restoration when a backup exists. The Rollback command remains available for the last successful backed-up operation. This is file restoration, not an in-memory ETABS transaction.

## Column and source-area detection

The selection reader accepts ETABS frame objects and reports all rejected objects in the DataGrid. It reads both frame endpoints, compares their global Z coordinates, and treats the higher endpoint as the column head. A frame is accepted only when:

```text
abs(delta Z) > horizontal length * vertical ratio tolerance
```

The reader first obtains area connectivity from the top joint. It filters connected objects by ETABS area object type and later validates elevation, opening state, and slab/deck property type. The fallback scans same-elevation areas and uses a tolerance-aware point-in-polygon/boundary test at the column head. Drop-envelope intersection is also considered so one requested drop can cross into adjacent source areas that do not own the column-head point. The implementation has no four-area assumption.

Drop dimensions are interpreted in the ETABS present coordinate units displayed in the window. A present-unit change after Preview blocks Apply.

## Geometry processing

`NetTopologySuite` version `2.5.0` is referenced only by the Application project. Its [NuGet package metadata](https://www.nuget.org/packages/NetTopologySuite/2.5.0) identifies `netstandard2.0`, and the package compiles in the repository's .NET Framework 4.8 solution.

For every selected column, the processor creates a centered rectangle in Global X-Y, column-local rotation, or a user-defined rotation. It then processes each source area once:

```text
DropPart = SourceArea intersection UnaryUnion(relevant drop polygons)
NormalPart = SourceArea difference DropPart
```

Same-elevation openings are subtracted from both results. Geometry is fixed after Boolean operations, snapped through a precision model, stripped of repeated and unnecessary collinear points, and filtered by minimum area. Polygons with holes are triangulated into simple hole-free polygons before ETABS creation. Invalid input rings, self-intersections, collapsed rings, slivers, and incomplete source mappings invalidate the complete plan.

Each generated region retains its `SourceAreaName` and `DropPanelAreaAssignmentBackup`; source mapping is never inferred after polygonization. Normal regions keep the source property. Drop regions override only the area property.

Adjacent polygons are merged only inside a homogeneous group containing one source area, one result type, and one deterministic complete assignment signature. The SHA-256 signature includes the resulting property, direct loads and their restore behavior, Shell Uniform Load Set assignments, local axis, local 3 direction, diaphragm, mesh rows, modifiers, groups, and pier/spandrel labels. Regions from different source areas are never merged.

## Assignment preservation

Before deletion, each source DTO contains:

- area name, label, story, property, and opening state;
- original winding and global local-3 direction;
- local-axis angle and advanced-axis flag;
- every direct area uniform load;
- Shell Uniform Load Set assignment names;
- diaphragm, modifiers, groups, pier, and spandrel labels;
- raw editable mesh-assignment table rows and schema.

Direct loads are keyed and verified by load pattern, load type, coordinate system, direction, and tolerance-compared value. `AreaObj.GetLoadUniform` does not return the setter's `Replace` argument, so the backup derives a deterministic restore instruction: the first load in each load pattern replaces that pattern's existing collection and later loads append. This prevents a later setter call from removing an earlier restored load.

Shell Uniform Load Set **assignments** are read and written through the ETABS table `Area Load Assignments - Uniform Load Sets`. They are independent of load-set definitions and direct uniform loads. Table object identity supports both unique-name and label/story schemas.

The referenced `ETABSv1` assembly does not expose `AreaObj.GetAutoMesh`. The service therefore resolves an importable area mesh-assignment table by known aliases and table metadata, preserves its raw rows, and verifies non-identity values after restoration. If no editable mesh table is available, Preview stops while **Preserve Mesh Assignments** is enabled. It never pretends that mesh preservation succeeded.

Advanced local-axis definitions cannot be restored through the angle-only `AreaObj.SetLocalAxes` method in this interop version. If a source uses advanced axes and local-axis preservation is enabled, Preview displays the validation error and blocks Apply.

Before `AreaObj.AddByCoord`, region winding is reversed when necessary to align the new polygon normal with the source local-3 direction. The area is then created with its final property, the local-axis angle and remaining assignments are restored, and local 3 is read back from the transformation matrix.

## ETABS API calls

Signatures and return behavior were inspected from the repository's `ETABSv1.dll`. The adapter checks the integer return code of every invoked method that exposes one.

| Purpose | Referenced ETABS API |
| --- | --- |
| Model context | `cSapModel.GetVersion`, `GetModelFilename`, `GetModelIsLocked`, `GetPresentUnits` |
| Selection | `SelectObj.GetSelected`, `SelectObj.ClearSelection` |
| Column endpoints and data | `FrameObj.GetPoints`, `GetSection`, `GetLabelFromName`, `GetTransformationMatrix` |
| Point data and connectivity | `PointObj.GetCoordCartesian`, `GetConnectivity` |
| Area discovery and geometry | `AreaObj.GetNameList`, `GetPoints`, `GetProperty`, `GetLabelFromName`, `GetOpening` |
| Slab filtering | `PropArea.GetSlab`, `PropArea.GetDeck` |
| Direct loads | `AreaObj.GetLoadUniform`, `SetLoadUniform` |
| Local axes and local 3 | `AreaObj.GetLocalAxes`, `SetLocalAxes`, `GetTransformationMatrix` |
| Other assignments | `AreaObj.Get/SetDiaphragm`, `Get/SetModifiers`, `GetGroupAssign`, `SetGroupAssign`, `Get/SetPier`, `Get/SetSpandrel` |
| Table assignments | `DatabaseTables.GetAvailableTables`, `GetTableForDisplayArray`, `GetTableForEditingArray`, `SetTableForEditingArray`, `ApplyEditedTables`, `CancelTableEditing` |
| Mutation | `AreaObj.Delete`, `AreaObj.AddByCoord` |
| Backup and rollback | `File.Save`, `File.OpenFile` |
| Highlight and refresh | `AreaObj.SetSelected`, `View.RefreshView` |

## Verification and logging

When verification is enabled, every created area is read back and compared with the expected source mapping. Checks cover:

- final area property;
- direct area loads;
- Shell Uniform Load Set assignments;
- local-axis angle and local-3 direction;
- diaphragm and mesh table rows;
- modifiers and group assignments;
- pier and spandrel labels.

Floating-point values use tolerance comparisons. Every mismatch records source area, new area, assignment type, expected value, actual value, and error message. Failed areas are selected in ETABS when possible. The UI reports verification failure distinctly and keeps Rollback available; it does not display a successful verification message.

Diagnostics include startup, attachment, selected/affected counts, validation, backup path, source and created names, restoration/verification results, rollback, and exceptions. High-risk ETABS mutation also uses the shared `CsiOperationLogger`.

**Export Log** writes a two-dimensional array to `DROP_PANEL_LOG` in one range assignment. It preserves and restores `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation`, and the active worksheet. Transient Excel ranges and worksheet collection wrappers are released; the shared Excel application and workbook are not released.

## Settings

`DropPanelSettingsStore` serializes `DropPanelOptions` into the user-scoped `DropPanelSettingsXml` value in the existing `Properties.Settings` system. It persists the selected property, dimensions, rotation, tolerances, all preservation flags, backup, merge, and verification options.

## Three-column irregular-layout test

`DropPanelGeometryProcessorTests.BuildPlan_processes_three_columns_and_preserves_source_mapping_across_irregular_slabs` models:

- three column heads on one story;
- three irregular source slabs;
- a drop from column C1 crossing S1 and S2;
- source S1 affected by C1 and C2;
- an opening near and inside the requested drop geometry;
- different local-axis angles, direct-load patterns, and Shell Uniform Load Set assignments.

The assertions prove that each source is processed once, source mapping survives polygon conversion, drop and normal properties are correct, the opening remains empty, output polygons are simple and valid, and different assignment signatures remain distinct. Rotation-mode behavior is covered separately.

These are deterministic pure-geometry tests. A final integration check against the installed ETABS build and representative `.edb` models is still required before production release because this development environment cannot start or mutate the user's live ETABS model.

## Build and deployment

Build with Visual Studio 2022 and the existing solution configuration, or run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  ExcelCSIToolBox.sln /t:Build /p:Configuration=Debug /p:Platform="Any CPU"
```

Run the tests with:

```powershell
dotnet test tests/ExcelCSIToolBox.Tests/ExcelCSIToolBox.Tests.csproj --configuration Debug
```

The existing deployment mechanism is unchanged. Install or update `bin/Debug/ExcelCSIToolBoxAddIn.vsto` through the normal Office customization installer flow. The target framework, language version, VSTO project type, assembly signing, and deployment manifests were not changed. `NetTopologySuite.dll` is copied into the normal add-in output and is included by the existing project-reference/deployment build.
