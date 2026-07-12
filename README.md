# Excel CSI ToolBox Add-In

This repository contains a VSTO add-in for Microsoft Excel that integrates Excel workflows with CSI products such as ETABS and SAP2000.

The add-in adds an Excel ribbon tab with tool windows for ETABS and SAP2000, allows attaching to a running CSI model, reads/writes data via Excel ranges, and includes utilities for post-processing results.

## Key Features

- Connects Excel to ETABS and SAP2000 using the CSI Open API.
- Launch ETABS/SAP2000 toolboxes directly from the Excel ribbon.
- Import/export result tables such as base reactions, modal mass participation ratios, story forces, story drifts, and mass summaries by story.
- Create and update model objects from Excel data.
- WPF-based tool windows following MVVM patterns.
- Experimental AI/MCP layer for assistant-style interactions with models.

## Solution Structure

```
ExcelCSIToolBox.sln
|
+-- src/
|   +-- ExcelCSIToolBoxAddIn              Main Excel VSTO add-in project
|   +-- ExcelCSIToolBox.Application       Use cases and workflow orchestration
|   +-- ExcelCSIToolBox.Core              Shared contracts, results, common logic
|   +-- ExcelCSIToolBox.Data              DTOs and mapper models during migration
|   +-- ExcelCSIToolBox.Infrastructure    ETABS/SAP2000 API adapters, Excel interop
|   +-- ExcelCSIToolBox.AI                AI/chatbox/MCP integration layer
+-- tests/
|   +-- ExcelCSIToolBox.Tests             Unit tests
+-- tools/
    +-- ExcelCSIToolBox.RefBuilder        CSI API reference/index utility
```

## Requirements

- Windows
- Microsoft Excel (desktop)
- .NET Framework 4.8
- Visual Studio Tools for Office (VSTO) Runtime
- ETABS and/or SAP2000 if using CSI-connected features
- CSI interop assemblies in `lib/` (for example `ETABSv1.dll`, `SAP2000v1.dll`)

If you install via `publish/ExcelCSIToolbox/setup.exe`, the installer can include the .NET Framework 4.8 and VSTO Runtime if missing.

## Installing (User)

Follow these steps to install the add-in for regular use (no development required):

1. Clone the repo or download the ZIP:

```powershell
git clone <repo-url>
```

2. If you downloaded a ZIP, right-click the ZIP file, open **Properties**, select **Unblock** if present, then extract. This prevents Windows blocking the `.vsto`/installer files after extraction.

3. Close all running Excel instances.

4. Open the publish folder:

```
publish/ExcelCSIToolbox/
```

5. Run:

```
setup.exe
```

6. If Windows warns about a temporary/self-signed certificate, proceed only if you trust the repository source.

7. Open Excel. You should see a new ribbon tab named **ExcelCSIToolBox**.

8. Use **ETABS Toolbox** or other tools. For CSI-connected features, open ETABS/SAP2000 and a model first, then attach from the toolbox.

## If the Add-in Does Not Appear

- Go to **File > Options > Add-ins** in Excel.
- At the bottom, in **Manage**, select **COM Add-ins** and click **Go...**.
- Ensure `ExcelCSIToolBoxAddIn` is checked.
- If the add-in is listed under **Disabled Items**, re-enable it via **Manage: Disabled Items**.
- If issues persist, close Excel and run `publish/ExcelCSIToolbox/setup.exe` again.

## Uninstall

1. Close Excel.
2. Open **Settings > Apps > Installed apps**.
3. Find `ExcelCSIToolBoxAddIn` and choose **Uninstall**.

## Build and Debug (Developer)

Follow these steps to build and debug from source:

1. Install Visual Studio with the **Office/SharePoint development** workload.
2. Ensure the .NET Framework 4.8 Developer Pack is installed.
3. Open the solution:

```
ExcelCSIToolBox.sln
```

4. Restore and build the solution in Visual Studio.
5. Set `ExcelCSIToolBoxAddIn` as the startup project.
6. Start debugging. Visual Studio will launch Excel with the add-in loaded.
7. Start ETABS/SAP2000 and open a model before testing features that rely on the CSI API.

## Publishing the Installer

To create a new installer:

1. Open `src/ExcelCSIToolBoxAddIn` in Visual Studio.
2. Choose **Build > Publish ExcelCSIToolBoxAddIn**.
3. Publish to:

```
publish/ExcelCSIToolbox/
```

4. Verify the publish folder contains `setup.exe`, `ExcelCSIToolBoxAddIn.vsto`, and an `Application Files/` folder.

5. Distribute the `publish/ExcelCSIToolbox/` folder to users for installation via `setup.exe`.

## Troubleshooting

- Add-in not visible: check COM Add-ins and Disabled Items in Excel.
- Certificate/trust errors: published files may use a temporary certificate. Re-publish using a trusted certificate if needed.
- Windows blocking `.vsto` files: Unblock the ZIP before extraction or unblock the installer files individually.
- Missing VSTO Runtime: run `setup.exe` rather than trying to load the `.vsto` directly.
- CSI connection issues: ensure ETABS/SAP2000 and a compatible model are open and the interop DLL versions in `lib/` match your CSI product versions.

## Notes for Contributors

- Target framework: **.NET Framework 4.8**.
- Host: Microsoft Excel via VSTO.
- UI: WPF with MVVM-style ViewModels.
- Keep CSI API access isolated inside Infrastructure adapters.
- Keep UI code lightweight; place workflow logic in Application/Core projects.
- `RefBuilder` is a utility used for generating reference scaffolding and is not part of the runtime add-in flow.
