# Naming Convention

Names should expose ownership and behavior. Prefer precise nouns and verbs over vague utility names.

## Projects and folders

- Projects use `ExcelCSIToolBox.<Boundary>` except the existing VSTO host `ExcelCSIToolBoxAddIn`.
- Feature folders use product or user language: `AnalysisResults`, `Connectivity`, `Loadings`, `Sections`, `ModellingHelpers`.
- Infrastructure folders are product first, capability second: `CSI/Etabs/DatabaseTables`, `CSI/Sap2000/Session`, `Excel/Writing`.

## Types

- Interfaces start with `I` and describe capability: `ISelectedObjectIdentityResolver`.
- Use cases end with `UseCase`: `GetFrameSectionsUseCase`.
- Requests and results end with `Request` and `Result`.
- DTOs end with `DTO` when they represent transport/API-shaped data.
- ViewModels end with `ViewModel`; WPF views end with `View` or `Window`; WinForms types end with `Form`.
- Async methods end with `Async`.
- Boolean properties should read as predicates: `IsBusy`, `CanApply`, `HasSelection`.
- ETABS-specific classes start with `Etabs`; SAP2000-specific classes start with `Sap2000`; Excel-specific classes start with `Excel` when they own Excel behavior.

Avoid vague names:

```text
Manager, Processor, Helper, Utility, CommonService, DataService, NewService, Handler2, Misc
```

Exceptions require a clear existing pattern or a narrowly scoped technical meaning.

Correct:

```text
EtabsDatabaseTableService
ExcelOutputService
ExportSelectedObjectConnectivityUseCase
BatchProgressReporter
```

Incorrect:

```text
CsiHelper
ExcelManager
DataService2
```

## When to update this document

Update this document when new project names, module names, generated file patterns, XML profile names, or resource naming conventions are introduced.

## Related documents

- [repository-structure.md](repository-structure.md)
- [coding-convention.md](coding-convention.md)
- [git-convention.md](git-convention.md)

Checklist:

- Name reveals boundary and capability.
- Product-specific types include product prefix.
- Tests mirror subject names.
- Vague suffixes are avoided.

