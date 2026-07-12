# CSI RefBuilder

`ExcelCSIToolBox.RefBuilder` is a development-time utility. It is not used by the AI Agent or MCP server at runtime.

Run from the repository root:

```powershell
dotnet run --project tools\ExcelCSIToolBox.RefBuilder\ExcelCSIToolBox.RefBuilder.csproj -- .
```

The pipeline writes API indexes to:

- `docs/reference/csi/ETABS/index/etabs_api_index.json`
- `docs/reference/csi/SAP2000/index/sap2000_api_index.json`

It also writes generated review summaries beside the current Infrastructure services. The runtime flow remains:

User chat -> AI Agent -> MCP Tool -> Infrastructure/Core service -> CSI API through `cSapModel`.
