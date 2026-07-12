# ViewModel Convention

ViewModels coordinate UI state and Application use cases. They must not become service locators, COM wrappers, or product-specific API adapters.

## Mandatory rules

- No direct `ETABSv1`, `SAP2000v1`, or Excel Interop types in ViewModels.
- Do not construct concrete Infrastructure services inside feature ViewModels.
- Constructor dependencies should be explicit and limited to the feature workflow.
- Initialize commands in the constructor or a dedicated initialization method.
- Expose busy, error, validation, and selected-item state clearly.
- Clean up event subscriptions and disposable dependencies.
- Use `ObservableCollection<T>` for bindable mutable collections and plain read-only models for immutable state.
- Split large shell ViewModels by existing partial feature files only when they are already part of the shell pattern. Do not create new partials for unrelated logic.

Feature ViewModel template:

```text
Fields: dependencies and backing state
Constructor: guard dependencies and initialize commands
Bindable properties: input, output, validation, busy/error state
Commands: execute/can execute
Private methods: validate, map UI input, call use case, update state
Cleanup: unsubscribe/dispose when needed
```

Correct:

```csharp
public ExportConnectivityViewModel(ExportSelectedObjectConnectivityUseCase useCase)
```

Incorrect:

```csharp
public ExportConnectivityViewModel()
{
    _service = new EtabsDatabaseTableService();
}
```

## When to update this document

Update this document when ViewModel base classes, command patterns, dialog abstractions, or shell ViewModel composition changes.

## Related documents

- [ui-convention.md](ui-convention.md)
- [application-logic-convention.md](application-logic-convention.md)
- [dependency-injection-convention.md](dependency-injection-convention.md)

Checklist:

- ViewModel has no direct COM references.
- Dependencies are injected or provided by the composition root.
- Busy/error/validation state is represented.
- Event subscriptions are cleaned up.

