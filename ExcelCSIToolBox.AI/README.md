# ExcelCSIToolBox.AI

- This project is reserved for future offline AI/MCP/Ollama integration.
- No production AI logic is implemented yet.
- Future MCP server should run locally/offline.
- Future MCP tools should call services from ExcelCSIToolBox.Application or ExcelCSIToolBox.Infrastructure through clean abstractions.
- AI/MCP must not reference UI/ViewModels/Views/Ribbon code.
- AI/MCP must not expose raw SapModel directly.
- Tool outputs should be structured and serialization-friendly.
- Transport should be abstracted so stdio, local HTTP, or named pipe can be swapped later.
- Product-specific ETABS/SAP2000 logic belongs in Infrastructure adapters/services, not in AI/MCP core classes.
