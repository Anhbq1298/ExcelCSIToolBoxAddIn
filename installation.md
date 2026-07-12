# Developer Installation & Build Guide

This guide describes how to clone the repository, build the project locally, and install the Excel VSTO Add-in directly from the compiled `bin/Debug` folder.

---

## 🛠️ Prerequisites
Before building and installing the add-in, ensure your computer has the following tools installed:
1. **Operating System**: Windows 10 or Windows 11.
2. **Microsoft Excel**: Office 2016, 2019, 2021, or Microsoft 365 (Desktop Edition).
3. **IDE**: [Visual Studio 2022](https://visualstudio.microsoft.com/vs/) (Community, Professional, or Enterprise).
   - During Visual Studio installation, make sure the **Office/SharePoint development** workload is checked.
4. **.NET Framework 4.8 Developer Pack**: [Download link](https://dotnet.microsoft.com/download/dotnet-framework/net48).
5. **VSTO Runtime**: [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=105522).

---

## 🚀 Build & Installation Steps

### Step 1: Clone the Repository
Open your terminal (PowerShell, Command Prompt, or Git Bash) and run the following command to clone the repository:
```bash
git clone https://github.com/Anhbq1298/ExcelCSIToolBoxAddIn.git
```
Navigate into the project directory:
```bash
cd ExcelCSIToolBoxAddIn
```

### Step 2: Open and Build the Solution

#### Option A: Using Visual Studio (Recommended)
1. Open the solution file `ExcelCSIToolBox.sln` in Visual Studio 2022.
2. In the top toolbar, ensure the build configuration is set to **Debug** and the platform is set to **Any CPU** (or **Active**).
3. Build the solution by selecting **Build** -> **Build Solution** from the top menu, or press `Ctrl + Shift + B`.

#### Option B: Using Command Line (MSBuild)
Alternatively, you can build the solution using MSBuild via PowerShell:
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" ExcelCSIToolBox.sln /t:Build /p:Configuration=Debug
```

---

### Step 3: Install the Add-in from the Debug Folder
Once the build completes successfully:
1. Open the compiled output folder:
   - Path: `[Cloned_Repository_Root]\bin\Debug\`
2. Double-click the deployment manifest file named **`ExcelCSIToolBoxAddIn.vsto`** inside this folder.
3. The **Microsoft Office Customization Installer** dialog will appear. Click **Install**.
4. Once completed, restart **Microsoft Excel**. You will see the new **CSI Toolbox** tab on the Excel Ribbon bar.

*Note: Since the add-in is registered directly from this `bin/Debug` folder, any subsequent code modifications and rebuilds will be automatically loaded by Excel without needing to reinstall the add-in again.*

---

## 🛑 Troubleshooting

### 1. CSI Toolbox Tab is Missing in Excel
If the tab does not show up after installation:
1. Go to Excel: **File** -> **Options** -> **Add-ins**.
2. At the bottom, change the **Manage** dropdown to **COM Add-ins** and click **Go...**
3. Locate **`ExcelCSIToolBoxAddIn`**:
   - If unchecked, **check** the box.
   - Check the **Load Behavior** message at the bottom:
     - If it says `Load at Startup` but is not loaded, proceed to the step below.

### 2. Trust & Security Blocks (Runtime Error)
If Excel blocks the add-in from loading due to trust policies:
1. Open Excel, go to **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...**
2. Click on **Add-ins** on the left column:
   - **Uncheck** the box `Require Application Add-ins to be signed by a Trusted Publisher`.
3. Click on **Macro Settings**:
   - Select `Disable VBA macros with notification` or `Enable all macros`.
4. Restart Excel to apply the changes.

---

## 🗑️ How to Uninstall
If you want to remove the add-in or perform a clean reinstall:
1. Close Microsoft Excel.
2. Open Windows **Settings** -> **Apps** -> **Installed Apps** (or **Control Panel** -> **Uninstall a Program**).
3. Search for **`ExcelCSIToolBoxAddIn`**.
4. Click **Uninstall** and follow the prompts.
