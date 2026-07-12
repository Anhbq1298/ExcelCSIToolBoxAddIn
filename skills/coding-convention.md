# Coding Convention

This repository targets .NET Framework 4.8 and C# 7.3 in the SDK-style projects. Do not use language features that require newer C# versions, such as file-scoped namespaces, records, init-only setters, or nullable reference type annotations.

## Mandatory rules

- Use one primary class per file.
- Use block-scoped namespaces.
- Prefer explicit types for public-facing code and complex LINQ; use `var` when the right side makes the type obvious.
- Use guard clauses for invalid constructor arguments and invalid public method inputs.
- Name async methods with `Async` and pass `CancellationToken` through async workflows when available.
- Do not call CSI COM APIs from thread-pool code such as `Task.Run`.
- Do not retain COM objects in DTOs, ViewModels, or Application contracts.
- Comments and XML documentation must be in English.
- Keep comments useful: explain why, not what obvious code does.
- Prefer `IReadOnlyList<T>` and `IEnumerable<T>` for read-only outputs; use concrete collections when mutation is part of the contract.
- Keep exceptions for exceptional conditions; use `OperationResult` for expected validation/API failures.

## Guidance

Methods should stay focused enough to review in one screen when practical. Larger methods are acceptable in generated code and legacy parsing/orchestration code, but new logic should be decomposed by responsibility. Partial classes are allowed for WPF/WinForms designer code and very large ViewModels already split by feature module; do not use partial classes to hide unrelated responsibilities.

Use LINQ for clear filtering and projection. Avoid dense LINQ chains when imperative code makes error handling or COM cleanup clearer.

Use `IDisposable` when a type owns an external resource, event subscription, unit scope, or temporary model state. Always document restore or cleanup failure behavior.

Correct:

```csharp
public ExportSelectedObjectConnectivityUseCase(
    IEtabsDatabaseTableService tableService,
    ISelectedObjectIdentityResolver identityResolver)
{
    _tableService = tableService ?? throw new ArgumentNullException(nameof(tableService));
    _identityResolver = identityResolver ?? throw new ArgumentNullException(nameof(identityResolver));
}
```

Incorrect:

```csharp
public void Run()
{
    Task.Run(() => _sapModel.FrameObj.GetNameList(ref count, ref names));
}
```

## When to update this document

Update this document when target frameworks, language versions, async patterns, COM lifetime rules, or result/exception policies change.

## Related documents

- [error-handling-convention.md](error-handling-convention.md)
- [csi-api-convention.md](csi-api-convention.md)
- [excel-interop-convention.md](excel-interop-convention.md)

Checklist:

- Code compiles with C# 7.3.
- Public APIs are clear and guarded.
- COM lifetime is explicit.
- Formatting-only churn is avoided.

