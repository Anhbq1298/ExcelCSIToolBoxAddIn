# Coding Guidelines & Rules for ExcelCSIToolBoxAddIn

## ETABS Open Application Programming Interface (OAPI)

### 1. Selection Retrieval & Identifier Resolution
- When querying active selection using `sapModel.SelectObj.GetSelected`, the ETABS OAPI returns numeric **Unique Names** (e.g. `"1"`, `"2"`).
- However, standard database tables (such as `"Element Forces - Columns"`, `"Element Forces - Beams"`, `"Element Forces - Braces"`, `"Pier Forces"`) reference elements by their user-friendly **Labels** (e.g. `"C1"`, `"B1"`, `"P1"`).
- **Rule**: When building or modifying selection filtering, always resolve **both the Unique Name and the Label** for selected elements (using `GetLabelFromName`) and match against both.
- Refer to [ETABS_API_GUIDELINE.md](file:///c:/repo/ExcelCSIToolBoxAddIn/ETABS_API_GUIDELINE.md) in the repository root for the full guide and candidate field keys catalog.

### 2. Table Column Headers
- Standard ETABS tables do not have standard column names like `"UniqueName"` or `"Frame"`. Columns will be named after the structural element type (e.g., `"Column"`, `"Beam"`, `"Brace"`, `"Shell"`).
- **Rule**: Always include these element-specific keys in candidate list matching arrays when calling `FindFieldIndex` to avoid missing the column index.
