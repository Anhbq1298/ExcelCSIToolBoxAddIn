# Application Logic Convention

Application code lives in `src/ExcelCSIToolBox.Application`. It expresses user workflows and coordinates Core contracts. It does not know about WPF, WinForms, VSTO, Excel Interop, `ETABSv1`, or `SAP2000v1`.

## Use-case structure

Use feature folders under `Application/Features`. A simple read workflow may only need a use case. A larger workflow should colocate request, result, validator, mapper, and use case files.

Templates:

```text
Features/<Feature>/<Workflow>Request.cs
Features/<Feature>/<Workflow>Result.cs
Features/<Feature>/<Workflow>Validator.cs
Features/<Feature>/<Workflow>UseCase.cs
Features/<Feature>/<Workflow>Mapper.cs
```

Read use case:

```text
Validate request -> call abstraction -> map result -> return OperationResult<T>
```

Write use case:

```text
Validate request -> capture model state/units -> perform checked API operation through abstraction -> restore state -> return OperationResult
```

Batch use case:

```text
Validate batch -> execute items in order -> collect item success/failure -> return partial success when some items fail
```

Export use case:

```text
Resolve selection -> read table -> filter rows -> prepare export model -> leave Excel writing to Infrastructure/AddIn
```

## Mandatory rules

- Depend on Core abstractions or Application interfaces, not concrete Infrastructure classes.
- Use `OperationResult` for validation, empty selection, partial failure, and external API failure.
- Do not expose COM types in request/result models.
- Capture and restore units or model state for workflows that temporarily change them.
- Treat read-only workflows and write workflows differently. Write workflows require clearer validation and user confirmation at the UI/agent boundary.
- Resolve selected CSI objects by unique name where possible, with label/story fallback handled deliberately.
- Table filtering must use schema-aware field aliases rather than hard-coded single headers.

Correct:

```text
ExportSelectedObjectConnectivityUseCase depends on IEtabsDatabaseTableService and ISelectedObjectIdentityResolver.
```

Incorrect:

```text
ExportSelectedObjectConnectivityUseCase creates Microsoft.Office.Interop.Excel.Range or ETABSv1.cSapModel.
```

## When to update this document

Update this document when use-case folder patterns, result semantics, validation order, or model-state rules change.

## Related documents

- [architecture-convention.md](architecture-convention.md)
- [csi-api-convention.md](csi-api-convention.md)
- [error-handling-convention.md](error-handling-convention.md)

Checklist:

- Request/result contracts are UI- and COM-free.
- Dependencies are abstractions.
- Validation happens before external calls.
- Partial failures are reported item by item.

