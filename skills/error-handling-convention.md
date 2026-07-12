# Error Handling Convention

Use `OperationResult` for expected outcomes: validation failure, empty selection, unsupported table, CSI return-code failure, Excel cancel, partial batch failure, and cleanup/restore failure. Use exceptions for programmer errors, invalid construction, and unexpected failures that cannot be represented safely.

## Mandatory rules

- User-facing messages must be specific and actionable.
- Diagnostic messages should include feature, product, operation, and return code when available.
- Do not swallow exceptions silently.
- Preserve exception details in logs when returning a friendly `OperationResult`.
- Distinguish success, validation failure, API failure, partial success, warning, and cleanup failure.
- Empty results are not automatically errors; decide by workflow.
- Cleanup failures, such as unit restoration failure, must be reported or logged.
- Batch operations should return item-level errors and a summary.

CSI return-code handling:

```text
Call API -> inspect return code -> map to OperationResult failure with operation name and return code
```

Excel failure handling:

```text
Validate workbook/worksheet/range -> treat cancel as expected -> return clear message for invalid target
```

Correct:

```text
"Failed to read Beam Object Connectivity from ETABS. CSI return code: 7."
```

Incorrect:

```text
"Something went wrong."
```

## When to update this document

Update this document when `OperationResult` semantics change, logging captures new diagnostic fields, or batch behavior changes.

## Related documents

- [logging-convention.md](logging-convention.md)
- [csi-api-convention.md](csi-api-convention.md)
- [excel-interop-convention.md](excel-interop-convention.md)

Checklist:

- Expected failures use `OperationResult`.
- Messages separate user action from diagnostics.
- Partial failures include item details.
- Cleanup failures are visible.

