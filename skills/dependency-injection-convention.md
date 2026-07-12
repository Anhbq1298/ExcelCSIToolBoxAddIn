# Dependency Injection Convention

Composition currently lives in the VSTO AddIn host, especially `src/ExcelCSIToolBoxAddIn/AddIn/AddInCompositionRoot.cs`, plus existing Application composition helpers such as `Application/Composition`.

## Mandatory rules

- Use constructor injection for Application services, ViewModels, and testable components.
- ViewModels must not create concrete Infrastructure services.
- Use factories only when runtime selection is required, such as ETABS versus SAP2000 or dynamic task pane/view creation.
- Session and dispatcher lifetimes must respect COM threading and product attachment state.
- Excel application access should be host-owned or Infrastructure-owned, not stored in Application.
- AI tool registration should use existing MCP tool registry/module patterns.
- Tests should replace dependencies with NSubstitute mocks or small fakes.

## Lifetimes

- Composition root: AddIn lifetime.
- CSI connection services: AddIn/session lifetime, with explicit reconnect behavior.
- Dispatcher: tied to the owning UI/COM thread.
- ViewModels: window/task-pane lifetime.
- Use cases: create per composition scope or per use when they are stateless.

Correct:

```csharp
var useCase = new GetFrameSectionsUseCase(_etabsConnectionService);
var viewModel = new CsiToolboxViewModel(useCase, otherDependencies);
```

Incorrect:

```csharp
public class CreateSectionViewModel
{
    private readonly EtabsConnectionService _service = new EtabsConnectionService();
}
```

## When to update this document

Update this document when introducing a DI container, changing AddIn composition, changing ViewModel factories, or changing AI tool registration.

## Related documents

- [architecture-convention.md](architecture-convention.md)
- [viewmodel-convention.md](viewmodel-convention.md)
- [feature-development-guide.md](feature-development-guide.md)

Checklist:

- Dependencies are explicit.
- Runtime factories have a clear reason.
- COM lifetimes remain host/infrastructure owned.
- Tests can replace services.

