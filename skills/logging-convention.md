# Logging Convention

Logging is for diagnostics, auditability of model operations, and support. User messages and log messages are related but not identical.

## Mandatory fields when available

- Operation name, such as `ExportSelectedObjectConnectivity`.
- Product, such as ETABS, SAP2000, or Excel.
- Feature/module name.
- Model file/path when safe.
- Correlation ID for multi-step or batch operations.
- CSI return code or Excel failure context.
- Duration for slow operations.
- Batch counts: total, succeeded, failed, skipped.

## Levels

- Debug: detailed decisions and intermediate values that are safe to record.
- Info: operation started/completed and summaries.
- Warning: recoverable partial failures, cleanup issues, missing optional data.
- Error: failed operation, exception, external API failure.

Do not log sensitive workbook data, large table payloads, secrets, PFX passwords, or unbounded generated output. Avoid log growth by keeping summaries bounded and rotating files when file logging is active.

Correct:

```text
Info ExportSelectedObjectConnectivity ETABS Model=Tower.edb Rows=42 DurationMs=180
Warning CsiPresentUnitScope Restore failed ReturnCode=3
```

Incorrect:

```text
Debug Full worksheet dump: <thousands of cells>
```

## When to update this document

Update this document when log sinks, file locations, rotation, correlation IDs, or operation logging fields change.

## Related documents

- [error-handling-convention.md](error-handling-convention.md)
- [csi-api-convention.md](csi-api-convention.md)
- [release-convention.md](release-convention.md)

Checklist:

- Logs include operation/product context.
- User and diagnostic messages are separated.
- Large or sensitive data is not logged.
- Batch operations log summaries.

