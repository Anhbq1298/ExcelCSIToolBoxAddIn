# TableFormats
Parsed Excel/table input DTOs.
- Prefer `Infrastructure/Csi/TableFormats/` for table shapes shared by ETABS, SAP2000, and future CSI products.
- Keep these classes as simple property bags.
- Do not put CSI API calls, Excel interop, validation workflows, or UI state here.
- Use this folder for row/table shapes such as point Cartesian input, frame input, steel section input, and concrete section input.
- Do not store shared table shapes under product-specific folders such as `Etabs/` or `Sap2000/`.
