# CSI API Convention

CSI API implementations live in `src/ExcelCSIToolBox.Infrastructure/CSI`.

- Shared product-neutral implementation: `CSI/Common`
- ETABS implementation: `CSI/Etabs`
- SAP2000 implementation: `CSI/Sap2000`

## Mandatory rules

- Direct `ETABSv1` references must stay under `CSI/Etabs`.
- Direct `SAP2000v1` references must stay under `CSI/Sap2000`.
- Do not call CSI COM APIs from thread-pool threads.
- Use the established dispatcher/session boundary for API calls that must remain on the owning thread.
- Check every CSI return code and include operation context in failures.
- Validate connection state before reading or writing.
- Validate model lock state before write operations when the API requires it.
- Capture present units before changing them and restore them in a cleanup path.
- Prefer unique object names for selection and identity resolution; use label/story fallbacks only when the table requires them.
- Keep ETABS and SAP2000 differences explicit. Do not hide product-specific behavior behind a shared abstraction unless behavior is genuinely the same.

## Mandatory patterns

Attach to model:

```text
Resolve running product -> acquire model object -> validate model file/product -> return OperationResult
```

Read selected objects:

```text
Validate connection -> call product selection API -> resolve object type/name -> return Core identity model
```

Read database table:

```text
Validate table key -> call product database table API -> normalize headers/rows -> return table result
```

Edit database table or set model data:

```text
Validate request -> check lock/state -> capture units/state -> call API -> check return code -> refresh view if needed -> restore state
```

Run analysis and read results:

```text
Validate model path/save state -> run analysis with checked return code -> select result cases -> read arrays -> map to Application/Core models
```

## Error reporting

Report product, operation name, model context when available, and CSI return code. User-facing messages should be specific enough to act on; diagnostic details can go to logging.

Correct:

```text
Infrastructure/CSI/Etabs/DatabaseTables/EtabsDatabaseTableService.cs uses ETABSv1 and returns EtabsTableResult.
```

Incorrect:

```text
Application/Features/AnalysisResults reads ETABSv1.cSapModel directly.
```

## When to update this document

Update this document when adding a new CSI product, changing dispatcher behavior, changing return-code handling, or introducing a new table/edit workflow.

## Related documents

- [application-logic-convention.md](application-logic-convention.md)
- [error-handling-convention.md](error-handling-convention.md)
- [logging-convention.md](logging-convention.md)

Checklist:

- Direct COM usage is in the product folder.
- Return codes are checked.
- Units/model state are restored.
- ETABS/SAP2000 differences are visible.

