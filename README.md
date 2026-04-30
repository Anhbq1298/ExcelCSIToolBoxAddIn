# ExcelCSIToolBoxAddIn

ExcelCSIToolBoxAddIn is a Microsoft Excel VSTO add-in for integrating Excel-based engineering workflows with CSI products such as **ETABS** and **SAP2000**.

The add-in provides Excel ribbon commands that open dedicated toolbox windows for ETABS and SAP2000. After a running CSI application is attached and a `SapModel` object is obtained, the downstream workflow can share common logic for model interaction, Excel range I/O, and engineering data processing.

## Current scope

The repository is being structured as a multi-project C#/.NET Framework 4.8 solution with separate layers for UI, application logic, shared core logic, data/DTO models, infrastructure adapters, and AI-related tooling.

The current direction is:

- **Excel VSTO Add-in Shell**: Ribbon, add-in startup, and window launch logic.
- **WPF Toolbox UI**: ETABS and SAP2000 toolbox windows.
- **Application Layer**: Use-case orchestration and workflow-level logic.
- **Core Layer**: Shared result models, domain-neutral contracts, and DTO-free abstractions.
- **Data Layer**: DTOs, Excel mapping models, table schemas, and data structures.
- **Infrastructure Layer**: ETABS/SAP2000 API adapters, Excel interop services, COM/API integration, and external system access.
- **AI Layer**: Future chatbox AI, local LLM, Ollama, or MCP-related integration.

## Solution projects

```text
ExcelCSIToolBoxAddIn.sln
│
├── ExcelCSIToolBoxAddIn
│   └── Main Excel VSTO add-in project.
│
├── ExcelCSIToolBox.Application
│   └── Application use cases and workflow orchestration.
│
├── ExcelCSIToolBox.Core
│   └── Shared logic, contracts, results, and abstractions.
│
├── ExcelCSIToolBox.Data
│   └── DTOs, mapper models, Excel schema models, and data structures.
│
├── ExcelCSIToolBox.Infrastructure
│   └── ETABS, SAP2000, Excel interop, and external API implementations.
│
└── ExcelCSIToolBox.AI
    └── AI/chatbox-related integration layer.
```

## Folder directory

```text
ExcelCSIToolBoxAddIn/
│
├── ExcelCSIToolBoxAddIn.sln
├── ExcelCSIToolBoxAddIn.csproj
├── ExcelCSIToolBoxAddin.cs
├── ThisAddIn.Designer.cs
├── ThisAddIn.Designer.xml
├── ExcelCSIToolBoxAddInRibbon.cs
├── ExcelCSIToolBoxAddInRibbon.Designer.cs
├── ExcelCSIToolBoxAddInRibbon.resx
├── ExcelCSIToolBoxAddIn_TemporaryKey.pfx
│
├── AddIn/
│   └── WindowManager.cs
│
├── UI/
│   ├── Views/
│   │   ├── EtabsToolboxWindow.xaml
│   │   ├── EtabsToolboxWindow.xaml.cs
│   │   ├── Sap2000ToolboxWindow.xaml
│   │   ├── Sap2000ToolboxWindow.xaml.cs
│   │   ├── BatchProgressWindow.xaml
│   │   ├── BatchProgressWindow.xaml.cs
│   │   ├── LoadCombinationDetailsWindow.xaml
│   │   ├── LoadCombinationDetailsWindow.xaml.cs
│   │   ├── FrameSectionDetailWindow.xaml
│   │   └── FrameSectionDetailWindow.xaml.cs
│   │
│   ├── ViewModels/
│   │   ├── CsiToolboxViewModel.cs
│   │   ├── FrameSectionDetailViewModel.cs
│   │   ├── FrameSectionDimensionEditItem.cs
│   │   ├── SectionDimensionAnnotation.cs
│   │   └── ViewModelBase.cs
│   │
│   └── Helpers/
│       └── SectionShapeRenderer.cs
│
├── Properties/
│   ├── AssemblyInfo.cs
│   ├── Resources.resx
│   ├── Resources.Designer.cs
│   ├── Settings.settings
│   └── Settings.Designer.cs
│
├── icon/
│   ├── etabs.png
│   └── sap2000icon.jpg
│
├── _ref/
│   ├── CSI API ETABS v1.chm
│   └── CSI_OAPI_Documentation.chm
│
├── ExcelCSIToolBox.Application/
│   └── ExcelCSIToolBox.Application.csproj
│
├── ExcelCSIToolBox.Core/
│   └── ExcelCSIToolBox.Core.csproj
│
├── ExcelCSIToolBox.Data/
│   └── ExcelCSIToolBox.Data.csproj
│
├── ExcelCSIToolBox.Infrastructure/
│   └── ExcelCSIToolBox.Infrastructure.csproj
│
└── ExcelCSIToolBox.AI/
    └── ExcelCSIToolBox.AI.csproj
```

## Project reference map

RefBuilder is archived as a one-time utility and is not part of the active solution. The current clean direction is: AddIn references Core/Data/Application/Infrastructure/AI; Core has no project references; Application references Core/Data; Infrastructure references Core/Data/Application; AI references Core/Data/Application and does not reference Infrastructure.

```text
ExcelCSIToolBoxAddIn
├── ExcelCSIToolBox.Core
├── ExcelCSIToolBox.Data
├── ExcelCSIToolBox.Infrastructure
└── ExcelCSIToolBox.Application

ExcelCSIToolBox.Application
├── ExcelCSIToolBox.Core
└── ExcelCSIToolBox.Data

ExcelCSIToolBox.Core
└── ExcelCSIToolBox.Data

ExcelCSIToolBox.Infrastructure
├── ExcelCSIToolBox.Core
├── ExcelCSIToolBox.Data
├── ETABSv1.dll
└── SAP2000v1.dll

ExcelCSIToolBox.AI
├── ExcelCSIToolBox.Core
└── ExcelCSIToolBox.Data
```

## Architecture notes

The intended architecture is to keep the CSI-product-specific acquisition logic isolated inside adapters. ETABS and SAP2000 differ mainly in how the running application object is attached and how the initial `SapModel` is retrieved. Once `SapModel` is available, most downstream operations can be shared.

Current architectural direction:

```text
Excel Ribbon / WPF UI
        ↓
Application Use Cases
        ↓
Core Contracts / Results / Shared Logic
        ↓
Data DTOs / Excel Mapping Models
        ↓
Infrastructure Adapters
        ↓
ETABS API / SAP2000 API / Excel Interop
```

A future clean-up target is to reduce direct dependency pressure on `Core`. Ideally, `Core` should contain domain-neutral contracts, result models, and shared abstractions without needing to reference Data, Infrastructure, or UI directly.

## Prerequisites

- Windows with Microsoft Excel installed.
- Visual Studio with Office/SharePoint development workload.
- .NET Framework 4.8 developer tooling.
- Compatible CSI products depending on the workflow:
  - ETABS with `ETABSv1.dll`.
  - SAP2000 with `SAP2000v1.dll`.
- Microsoft Office interop assemblies / VSTO runtime.

## Build and run

1. Open `ExcelCSIToolBoxAddIn.sln` in Visual Studio.
2. Build the solution using the required configuration.
3. Start debugging from Visual Studio to launch Excel with the add-in loaded.
4. Use the custom Excel ribbon commands to open the ETABS or SAP2000 toolbox window.
5. Attach to a running ETABS/SAP2000 instance before running API-dependent operations.


## Notes for contributors

- Target framework: **.NET Framework 4.8**.
- UI pattern: WPF with MVVM-style ViewModels.
- Main host: Microsoft Excel through VSTO.
- CSI API access should be isolated behind Infrastructure adapters where possible.
- Shared workflows should depend on the common `SapModel` abstraction/usage pattern after the CSI model is acquired.
- Keep UI orchestration thin; place workflow logic in Application/Core where practical.

