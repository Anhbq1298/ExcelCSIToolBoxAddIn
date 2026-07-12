# Feature Development Guide

This guide shows the normal path for adding a feature. Example: add a new ETABS frame output export.

## Standard flow

1. Define Core contracts in `src/ExcelCSIToolBox.Core/Contracts` or models in `Core/Models` if new pure data is needed.
2. Define Application request/result under `src/ExcelCSIToolBox.Application/Features/AnalysisResults/<Workflow>`.
3. Create the use case and validator. Depend on abstractions, not concrete ETABS services.
4. Implement ETABS infrastructure under `src/ExcelCSIToolBox.Infrastructure/CSI/Etabs/AnalysisResults` or another exact capability folder.
5. Register dependencies in `src/ExcelCSIToolBoxAddIn/AddIn/AddInCompositionRoot.cs` or the existing composition bundle.
6. Create UI under `src/ExcelCSIToolBoxAddIn/UI/Modules/AnalysisResults`.
7. Create or update ViewModel commands without direct COM access.
8. Add navigation, Ribbon, or shell entry in the AddIn/UI module.
9. Add tests under `tests/ExcelCSIToolBox.Tests/Application/...`, `Infrastructure/...`, or `Architecture/...`.
10. Update skills/docs if the feature introduces a new convention.
11. Run restore, build, tests, architecture checks, and a manual smoke test when the required external apps are available.

## Flow variations

- Read-only ETABS feature: focus on selection/table/result mapping and no write confirmation.
- Write-to-model ETABS feature: add validation, model lock checks, preview/confirmation, return-code checks, and state restoration.
- SAP2000 feature: place implementation under `Infrastructure/CSI/Sap2000` and avoid reusing ETABS assumptions unless verified.
- Excel import feature: read Excel through `Infrastructure/Excel/Reading` or AddIn selection host, then pass plain values to Application.
- Excel export feature: Application prepares data, `Infrastructure/Excel/Writing` writes bulk values.
- Shared ETABS/SAP2000 feature: put only product-neutral logic under `CSI/Common`; keep product adapter calls separate.
- AI-exposed feature: add or update MCP tool metadata under `ExcelCSIToolBox.AI/Mcp/Tools` and route to Application/Infrastructure through existing tool context.

Correct:

```text
Analysis export use case in Application, ETABS table call in Infrastructure, WPF dialog in AddIn/UI.
```

Incorrect:

```text
WPF button handler calls ETABSv1 and writes directly to Excel cell by cell.
```

## When to update this document

Update this document when composition, module layout, test ownership, or feature registration changes.

## Related documents

- [repository-structure.md](repository-structure.md)
- [application-logic-convention.md](application-logic-convention.md)
- [testing-convention.md](testing-convention.md)

Checklist:

- The feature has one clear owner folder.
- COM is isolated.
- Tests mirror production ownership.
- Documentation is updated for new patterns.

