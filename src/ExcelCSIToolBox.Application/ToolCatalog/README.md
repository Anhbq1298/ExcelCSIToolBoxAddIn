# ToolCatalog

ToolCatalog contains metadata used by the AI planning layer to understand,
validate, and select available tools.

It includes:
- tool schema
- tool aliases
- required and optional parameters
- intent hints
- clarification messages
- validation rules

ToolCatalog does not execute tools and does not call ETABS/SAP2000 APIs.

Executable MCP tool wrappers live under:

`ExcelCSIToolBox.AI/Mcp/Tools`
