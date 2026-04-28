# ExcelCSIToolBox.AI

This project is reserved for future offline AI, MCP, Ollama, and chat orchestration work.

No production AI logic is implemented yet. The future MCP server should run locally/offline, with tool outputs kept structured and serialization-friendly. MCP tools should call application and infrastructure services through clean abstractions, never UI views, viewmodels, ribbon code, VSTO add-in classes, or raw SapModel objects.

Transport should remain abstract so stdio, local HTTP, or named pipes can be swapped later. Product-specific ETABS and SAP2000 logic belongs in Infrastructure adapters and services, not in AI/MCP core classes.
