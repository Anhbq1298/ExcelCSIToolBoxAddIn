# Column Drop Modelling Helper

## Overview

Column Drop is a modeless WPF modelling helper in the Excel VSTO add-in. It uses the ETABS connection already owned by CSIToolbox; the window does not attach to ETABS or display duplicated model metadata.

The window is available at:

```text
Excel -> ExcelCSIToolBox -> ETABS Toolbox -> MODELLING HELPER -> Helpers -> Drop Panel
```

## Workflow

1. Click **Select Columns**.
2. Select one or more frame objects in ETABS and confirm the selection in the shared interactive-selection window.
3. Enter **Drop Thickness** in the active ETABS model length unit.
4. Select an ETABS concrete material.
5. Click **Create Drop**.

The selection replaces the previous selection, removes duplicate frame names, and keeps only valid vertical columns. Cancelling the interactive selection leaves the existing in-memory selection unchanged and restores the modeless tool window.

The window displays only the selected column names, their count, the two user inputs, and the generated property-name preview. There is no preview/highlight workflow, ETABS attachment section, backup/rollback control, region-merge option, or post-apply verification control.

## Drop property

The tool reads the active ETABS length unit dynamically and interprets thickness in that unit. Concrete materials are read from the attached model.

The property name is generated deterministically as:

```text
Drop_<Thickness>_<MaterialName>
```

Unnecessary decimal zeros are removed and decimal values always use a dot in the property name. For example, thickness `1500` and material `C32/40` produce `Drop_1500_C32/40`.

At execution time, the material and current model unit are revalidated. A missing property is created as an ETABS `Drop` slab property with `ShellThick` behavior, the selected material, and the entered thickness. An existing property is reused only when its slab type, shell behavior, material, and thickness match. A same-name incompatible property stops the operation and is never overwritten.

## Shell selection and splitting

For each valid column, the higher frame endpoint is the column head. `PointObj.GetConnectivity` determines the area objects attached to that exact joint. Only horizontal slab/deck areas returned by that connectivity query and at the column-head elevation are eligible.

There is no fixed shell-count assumption. If four slab shells meet at one column head, the same drop boundary is intersected with all four shells:

```text
inside  = source shell intersection drop boundary
outside = source shell difference inside
```

Each affected source shell is processed independently. Inside polygons receive the generated/reused drop property. Outside polygons retain that source shell's original property. Regions from different source shells are not merged.

Same-elevation openings are subtracted from both inside and outside results. Invalid rings, collapsed polygons, and regions below the minimum-area tolerance stop the operation before ETABS mutation begins.

The referenced ETABS interop does not expose an area divide/cookie-cut API. The implementation therefore creates calculated replacement regions with `AreaObj.AddByCoord` and deletes each original source area after its replacement regions and API-based assignments have been created.

## Assignment handling

Every generated region keeps its source-area assignment backup. The tool preserves direct area loads, Shell Uniform Load Set assignments, local axes and local-3 orientation, diaphragm, mesh assignments, area modifiers, groups, and pier/spandrel labels. Only the area section property differs between inside and outside regions.

Shell Uniform Load Set and mesh assignments are restored through editable ETABS database tables where the referenced interop does not expose the required object methods. Advanced local-axis definitions that cannot be recreated safely are rejected before mutation.

## Mutation and reporting

Apply revalidates the active model, lock state, units, material, property definition, selected columns, source properties, and source coordinates. Replacement areas and ordinary API assignments are created before original source shells are deleted. The ETABS view is refreshed after a successful operation.

The success message reports processed columns, created drop objects, the created or reused property, thickness/unit, material, and skipped-column reasons. **Export Log** writes region records to `DROP_PANEL_LOG` in Excel.

## Tests

`DropPanelGeometryProcessorTests` covers irregular geometry, rotation modes, direct column-head connectivity, invalid source rings, openings, and four shells connected to one column head. `DropPanelPropertyNameBuilderTests` covers deterministic property names, invariant decimal formatting, and invalid input.

Build and test with:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  src\ExcelCSIToolBoxAddIn\ExcelCSIToolBoxAddIn.csproj /t:Build /p:Configuration=Debug /p:Platform="AnyCPU"

dotnet test tests\ExcelCSIToolBox.Tests\ExcelCSIToolBox.Tests.csproj --configuration Debug
```

An integration check against representative `.edb` models is still required because the automated tests do not mutate a live ETABS model.
