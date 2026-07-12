# Testing Convention

Automated tests live in `tests/ExcelCSIToolBox.Tests` and mirror production ownership:

```text
Architecture/
Application/
Core/
Infrastructure/
```

The current test stack is xUnit, FluentAssertions, and NSubstitute.

## Mandatory rules

- Name test classes after the subject, for example `GetFrameSectionsUseCaseTests`.
- Use Arrange/Act/Assert structure.
- Use `[Fact]` for one behavior and `[Theory]` for input matrices.
- Mock Core/Application abstractions, not COM objects.
- Do not directly unit-test ETABS/SAP2000/Excel COM calls in normal unit tests.
- Add regression tests before changing confirmed risky behavior.
- Add architecture tests for enforceable boundary rules.
- Geometry and table-schema logic should have deterministic tests.
- UI tests should be minimal unless a stable automation harness exists.

## Coverage expectations

Read workflows should test success, validation failure, empty result, and external failure. Write workflows should additionally test lock/state validation, return-code failure, partial failure, and cleanup/restoration behavior.

Manual smoke tests are still required for VSTO/Excel/ETABS/SAP2000 behavior when those applications are available:

```text
Excel starts with the add-in loaded.
ETABS attach/read/write path works.
SAP2000 attach/read/write path works.
Excel import/export writes expected cells.
```

Correct:

```text
tests/ExcelCSIToolBox.Tests/Application/Features/Sections/GetFrameSectionsUseCaseTests.cs
```

Incorrect:

```text
tests/ExcelCSIToolBox.Tests/MiscTests.cs directly instantiates ETABSv1.cSapModel.
```

## When to update this document

Update this document when test projects split, a new test library is introduced, or architecture checks change.

## Related documents

- [architecture-convention.md](architecture-convention.md)
- [application-logic-convention.md](application-logic-convention.md)
- [pull-request-checklist.md](pull-request-checklist.md)

Checklist:

- Tests sit near mirrored ownership.
- COM is replaced by abstractions or covered by manual smoke tests.
- Architecture tests pass.
- Test names describe behavior.

