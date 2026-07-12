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

## 🛠️ Prerequisites
Before building and installing the add-in, ensure your computer has the following tools installed:
1. **Operating System**: Windows 10 or Windows 11.
2. **Microsoft Excel**: Office 2016, 2019, 2021, or Microsoft 365 (Desktop Edition).
3. **IDE**: [Visual Studio 2022](https://visualstudio.microsoft.com/vs/) (Community, Professional, or Enterprise).
   - During Visual Studio installation, make sure the **Office/SharePoint development** workload is checked.
4. **.NET Framework 4.8 Developer Pack**: [Download link](https://dotnet.microsoft.com/download/dotnet-framework/net48).
5. **VSTO Runtime**: [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=105522).
6. **CSI Products**: ETABS and/or SAP2000 if using CSI-connected features.
7. **CSI interop assemblies** in `lib/` (for example `ETABSv1.dll`, `SAP2000v1.dll`).

---

## 🚀 Build & Installation Guide (Developer)

Follow these steps to clone, build, and install the Excel VSTO Add-in locally:

### Step 1: Clone the Repository
Open your terminal (PowerShell, Command Prompt, or Git Bash) and run:
```bash
git clone https://github.com/Anhbq1298/ExcelCSIToolBoxAddIn.git
cd ExcelCSIToolBoxAddIn
```

### Step 2: Open and Build the Solution

#### Option A: Using Visual Studio (Recommended)
1. Open the solution file `ExcelCSIToolBox.sln` in Visual Studio 2022.
2. In the top toolbar, ensure the build configuration is set to **Debug** and the platform is set to **Any CPU** (or **Active**).
3. Build the solution by selecting **Build** -> **Build Solution** from the top menu, or press `Ctrl + Shift + B`.

#### Option B: Using Command Line (MSBuild via PowerShell)
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" ExcelCSIToolBox.sln /t:Build /p:Configuration=Debug
```

### Step 3: Install the Add-in from the Debug Folder
Once the build completes successfully:
1. Open the compiled output folder:
   - Path: `[Cloned_Repository_Root]\bin\Debug\`
2. Double-click the deployment manifest file named **`ExcelCSIToolBoxAddIn.vsto`** inside this folder.
3. The **Microsoft Office Customization Installer** dialog will appear. Click **Install**.
4. Once completed, restart **Microsoft Excel**. You will see the new **CSI Toolbox** tab on the Excel Ribbon bar.

*Note: Since the add-in is registered directly from this `bin/Debug` folder, any subsequent code modifications and rebuilds will be automatically loaded by Excel without needing to reinstall the add-in again.*

---

## 📦 User Installation

To install the add-in for regular use (no development/build required):

1. Close all running Excel instances.
2. Open the publish folder: `publish/ExcelCSIToolbox/`
3. Run `setup.exe`.
4. If Windows warns about a temporary/self-signed certificate, proceed only if you trust the repository source.
5. Open Excel. You should see a new ribbon tab named **ExcelCSIToolBox**.

---

## 🛑 Troubleshooting & Common Issues

### 1. CSI Toolbox Tab is Missing in Excel
If the tab does not show up after installation:
1. Go to Excel: **File** -> **Options** -> **Add-ins**.
2. At the bottom, change the **Manage** dropdown to **COM Add-ins** and click **Go...**
3. Locate **`ExcelCSIToolBoxAddIn`**:
   - If unchecked, **check** the box.
   - Check the **Load Behavior** message at the bottom.
   - If the add-in is listed under **Disabled Items**, re-enable it via **Manage: Disabled Items**.

### 2. Trust & Security Blocks (Runtime Error)
If Excel blocks the add-in from loading due to trust policies:
1. Open Excel, go to **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...**
2. Click on **Add-ins** on the left column:
   - **Uncheck** the box `Require Application Add-ins to be signed by a Trusted Publisher`.
3. Click on **Macro Settings**:
   - Select `Disable VBA macros with notification` or `Enable all macros`.
4. Restart Excel to apply the changes.

### 3. Certificate/Trust or Zip Block Errors
- **Windows blocking `.vsto` files**: If downloaded as a ZIP, right-click the ZIP file, open **Properties**, select **Unblock** if present, then extract.
- **Missing VSTO Runtime**: Run `setup.exe` rather than trying to load the `.vsto` directly.
- **CSI connection issues**: Ensure ETABS/SAP2000 and a compatible model are open, and the interop DLL versions in `lib/` match your CSI product versions.

---

## 🗑️ How to Uninstall
If you want to remove the add-in or perform a clean reinstall:
1. Close Microsoft Excel.
2. Open Windows **Settings** -> **Apps** -> **Installed Apps** (or **Control Panel** -> **Uninstall a Program**).
3. Search for **`ExcelCSIToolBoxAddIn`**.
4. Click **Uninstall** and follow the prompts.

---

## 🚀 Publishing the Installer
To create a new installer:
1. Open `src/ExcelCSIToolBoxAddIn` in Visual Studio.
2. Choose **Build > Publish ExcelCSIToolBoxAddIn**.
3. Publish to `publish/ExcelCSIToolbox/`.
4. Verify the folder contains `setup.exe`, `ExcelCSIToolBoxAddIn.vsto`, and an `Application Files/` folder.
5. Distribute the `publish/ExcelCSIToolbox/` folder to users for installation via `setup.exe`.

## Notes for Contributors

- Target framework: **.NET Framework 4.8**.
- Host: Microsoft Excel via VSTO.
- UI: WPF with MVVM-style ViewModels.
- Keep CSI API access isolated inside Infrastructure adapters.
- Keep UI code lightweight; place workflow logic in Application/Core projects.
- `RefBuilder` is a utility used for generating reference scaffolding and is not part of the runtime add-in flow.
