# Coding Agent Instructions

This is the primary operating guide for future coding agents working in this repository.

## Mandatory pre-task checklist

1. Read this file first.
2. Read [architecture-convention.md](architecture-convention.md) and [repository-structure.md](repository-structure.md).
3. Read feature-specific documents for the task: CSI, Excel, UI, ViewModel, testing, or release.
4. Inspect the actual implementation before deciding placement.
5. Search for existing services, models, use cases, ViewModels, and helpers before adding new ones.
6. Check project references and current test coverage.
7. Build/test baseline when the task is structural or risky.

## Working rules

- Preserve existing behavior unless the user explicitly asks for behavior change.
- Respect project boundaries and feature-oriented folders.
- Avoid direct COM access outside Infrastructure or AddIn host boundaries.
- Never call CSI COM APIs from thread-pool threads.
- Do not create direct services inside ViewModels.
- Keep UI behavior and styling consistent with existing modules.
- Add or update tests for changed behavior and architecture checks for enforceable boundary rules.
- Update documentation when conventions change.
- Build after each meaningful phase.
- Report files changed, tests executed, and limitations.
- Never create vague dumping-ground folders.
- Never create duplicate abstractions without checking existing ones.
- Prefer extending established architecture over creating parallel architecture.
- Do not move files and change business logic in the same commit unless inseparable.
- Stop and report when the repository cannot build because of a pre-existing external dependency.

## Post-task checklist

- Files are in the correct project/folder.
- Namespaces and project references are valid.
- Resource paths and embedded resources were updated.
- Obsolete folder paths and linked compile items were searched.
- Direct `ETABSv1`, `SAP2000v1`, and Excel Interop references are confined.
- Automated tests and architecture tests were run.
- VSTO/Excel/ETABS/SAP2000 manual tests are reported only when actually run.
- Commits are small and meaningful.

Correct:

```text
Move UI files by module, update XAML resource paths, build tests, then commit refactor(ui): organize presentation by feature.
```

Incorrect:

```text
Move folders, rewrite business logic, skip tests, and claim Excel was verified without opening Excel.
```

## When to update this document

Update this document when agent workflow, repository boundaries, validation steps, or mandatory reporting requirements change.

## Related documents

- [architecture-convention.md](architecture-convention.md)
- [repository-structure.md](repository-structure.md)
- [testing-convention.md](testing-convention.md)

