# Excel Interop Convention

Excel COM code belongs in `src/ExcelCSIToolBox.Infrastructure/Excel` or in VSTO/AddIn host code that must interact with the active Excel UI. Application and Core must use plain contracts such as cell references, table export models, and value arrays.

## Mandatory rules

- Do not retain `Range`, `Worksheet`, `Workbook`, or `Application` objects in ViewModels or Application contracts.
- Prefer bulk `Value2` reads and writes over per-cell loops.
- Validate workbook, worksheet, active selection, and target anchor before reading or writing.
- Do not keep long-lived COM references unless VSTO owns the lifecycle.
- Release add-in-created COM objects when ownership is clear.
- Do not put formatting rules inside Application use cases.
- Use plain models such as `PreparedTableExport` and `ExcelCellReference` across non-Interop boundaries.
- Handle workbook-close and selection-cancel paths as expected user outcomes.

## Safe examples

```csharp
object[,] values = PreparedTableExportValueBuilder.BuildValues(export, includeHeaders: true);
// ExcelOutputService writes values to a target range in one operation.
```

```text
Application creates PreparedTableExport.
Infrastructure/Excel/Writing writes PreparedTableExport to Excel.
```

Unsafe examples:

```csharp
public class ExportResult
{
    public Microsoft.Office.Interop.Excel.Range TargetRange { get; set; }
}
```

```csharp
foreach (var cell in cells)
{
    worksheet.Cells[row, column].Value2 = cell;
}
```

## UI behavior

Excel range pickers should clearly handle cancel, invalid selection, merged cells, workbook close, and selection from a different workbook. Target anchor handling should not silently overwrite unrelated content without explicit workflow intent.

## When to update this document

Update this document when Excel reading/writing services change, new formatting behavior is added, or COM cleanup policy changes.

## Related documents

- [application-logic-convention.md](application-logic-convention.md)
- [ui-convention.md](ui-convention.md)
- [error-handling-convention.md](error-handling-convention.md)

Checklist:

- Bulk read/write is used where practical.
- COM objects do not cross into Core/Application.
- Workbook and selection state are validated.
- Cancel is not treated as an exception.

