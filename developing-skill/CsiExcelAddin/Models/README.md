# Models
DTOs and simple data models.
- Shared data transfer objects used across services and ViewModels.
- No business logic inside models.
- Parsed Excel/table input DTOs should be kept in a shared `Infrastructure/Csi/TableFormats/` folder when the same shape can be used by multiple CSI products.
- Examples of table-format DTOs: point Cartesian input, frame-by-coordinate input, frame-by-point input, steel section input, and concrete section input.
