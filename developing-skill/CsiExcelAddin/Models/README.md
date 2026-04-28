# Models
DTOs and simple data models.
- Shared data transfer objects used across services and ViewModels.
- No business logic inside models.
- Parsed Excel/table input DTOs should be kept in a dedicated `TableFormats/` folder when the project has one.
- Examples of table-format DTOs: point Cartesian input, frame-by-coordinate input, frame-by-point input, steel section input, and concrete section input.
