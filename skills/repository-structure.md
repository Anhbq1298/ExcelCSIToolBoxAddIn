# Repository Structure

The repository is organized by architectural boundary at the solution level and by feature or capability inside each project.

## Final source layout

```text
/
|-- src/
|   |-- ExcelCSIToolBox.Core/
|   |-- ExcelCSIToolBox.Application/
|   |-- ExcelCSIToolBox.Infrastructure/
|   |-- ExcelCSIToolBox.AI/
|   `-- ExcelCSIToolBoxAddIn/
|-- tests/
|   `-- ExcelCSIToolBox.Tests/
|       |-- Architecture/
|       |-- Application/
|       |-- Core/
|       `-- Infrastructure/
|-- tools/
|   `-- ExcelCSIToolBox.RefBuilder/
|-- docs/
|-- skills/
|-- build/
|-- lib/
`-- ExcelCSIToolBox.sln
```

The UI is currently inside the VSTO host project at `src/ExcelCSIToolBoxAddIn/UI` because this keeps Ribbon, task pane, WPF, and WinForms resources buildable with the existing AddIn project.

## Project purposes

- `ExcelCSIToolBox.Core`: pure abstractions, contracts, domain models, geometry, tabular primitives, and result types.
- `ExcelCSIToolBox.Application`: use cases, feature workflows, validators, mappers, and application services that depend on Core abstractions.
- `ExcelCSIToolBox.Infrastructure`: ETABS, SAP2000, Excel, and other external implementation details.
- `ExcelCSIToolBox.AI`: AI provider clients, agent orchestration, MCP contracts, MCP server/client code, and AI-exposed tool modules.
- `ExcelCSIToolBoxAddIn`: VSTO host, composition root, Ribbon/task panes, WPF/WinForms UI, and host-specific wiring.
- `ExcelCSIToolBox.Tests`: unit and architecture tests, mirrored by production ownership.
- `ExcelCSIToolBox.RefBuilder`: tooling for CSI reference/catalog generation.

## Placement decision table

| What does the file do? | Uses COM? | Depends on UI? | Product-specific? | Workflow? | Pure contract? | Place it here |
| --- | --- | --- | --- | --- | --- | --- |
| DTO, enum, plain model, result type | No | No | No | No | Yes | `src/ExcelCSIToolBox.Core` |
| Use-case request/result/validator | No | No | No | Yes | Maybe | `src/ExcelCSIToolBox.Application/Features/<Feature>` |
| ETABS implementation | ETABSv1 | No | ETABS | Maybe | No | `src/ExcelCSIToolBox.Infrastructure/CSI/Etabs/<Capability>` |
| SAP2000 implementation | SAP2000v1 | No | SAP2000 | Maybe | No | `src/ExcelCSIToolBox.Infrastructure/CSI/Sap2000/<Capability>` |
| Shared CSI implementation | Maybe hidden behind abstractions | No | Shared | Maybe | No | `src/ExcelCSIToolBox.Infrastructure/CSI/Common/<Capability>` |
| Excel Interop service | Excel Interop | No | Excel | Maybe | No | `src/ExcelCSIToolBox.Infrastructure/Excel/<Capability>` |
| WPF view, ViewModel, form, renderer | No direct CSI COM | Yes | UI | Maybe | No | `src/ExcelCSIToolBoxAddIn/UI/Modules/<Module>` or `UI/Shared` |
| VSTO lifecycle, task pane host, windows | VSTO or WPF host | Yes | AddIn | No | No | `src/ExcelCSIToolBoxAddIn/AddIn` |
| AI provider or MCP tool | No CSI COM directly | No | AI | Maybe | No | `src/ExcelCSIToolBox.AI` |

## Feature and shared folder rules

Feature folders should reveal a complete user workflow. For example, `Application/Features/Connectivity` contains connectivity export use cases, while `UI/Modules/AnalysisResults` contains the related user-facing components. Shared folders are allowed only when code is reused by more than one feature and the name describes a capability, such as `UI/Shared/Progress` or `Infrastructure/CSI/Common/Dispatching`.

Do not create generic dumping grounds such as `Misc`, `Helpers`, `NewFolder`, `Temp`, or `CommonService`.

Correct:

```text
src/ExcelCSIToolBox.Infrastructure/CSI/Etabs/DatabaseTables/EtabsDatabaseTableService.cs
src/ExcelCSIToolBoxAddIn/UI/Modules/Loadings/ShellUniformLoadSetForm.cs
```

Incorrect:

```text
src/ExcelCSIToolBox.Core/EtabsDatabaseTableService.cs
src/ExcelCSIToolBoxAddIn/UI/Helpers/ShellUniformLoadSetForm.cs
```

## When to update this document

Update this document when projects are added, responsibilities move between projects, feature folders are renamed, or architecture tests enforce a new placement rule.

## Related documents

- [architecture-convention.md](architecture-convention.md)
- [feature-development-guide.md](feature-development-guide.md)
- [naming-convention.md](naming-convention.md)

Checklist:

- New files are in the narrowest correct project.
- Feature-specific code is colocated.
- Shared folders are capability-specific.
- No linked cross-project compile items are introduced.

