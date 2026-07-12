# Pull Request Checklist

Use this checklist for every PR. Mark unavailable manual checks as not run and explain why.

## Build

- [ ] `dotnet restore .\ExcelCSIToolBox.sln`
- [ ] Core builds.
- [ ] Application builds.
- [ ] Infrastructure builds.
- [ ] AI builds.
- [ ] AddIn/VSTO builds, or limitation is documented.
- [ ] Tests build.

## Tests

- [ ] Automated tests pass.
- [ ] Architecture tests pass.
- [ ] Regression tests were added or updated for risky behavior.
- [ ] Manual ETABS smoke test was run when relevant.
- [ ] Manual SAP2000 smoke test was run when relevant.
- [ ] Manual Excel/VSTO smoke test was run when relevant.

## Architecture

- [ ] File placement matches [repository-structure.md](repository-structure.md).
- [ ] Namespaces match ownership or an intentional transition is documented.
- [ ] Core/Application do not reference UI, Infrastructure, or COM.
- [ ] ETABS API calls stay under `Infrastructure/CSI/Etabs`.
- [ ] SAP2000 API calls stay under `Infrastructure/CSI/Sap2000`.
- [ ] Excel Interop is isolated to Infrastructure/Excel or AddIn host/UI boundary.
- [ ] No linked compile items were introduced.
- [ ] No vague dumping-ground folders were introduced.

## UI and behavior

- [ ] Existing UI style is preserved.
- [ ] Screenshots are included where UI changed.
- [ ] Error, empty, busy, and validation states are handled.
- [ ] Write operations require confirmation or preview where appropriate.
- [ ] No unrelated workflow changes were included.

## Documentation

- [ ] Relevant skills/docs were updated.
- [ ] Public behavior changes are described.
- [ ] Known limitations are explicit.

Correct:

```text
PR notes: VSTO build not run because Microsoft.VisualStudio.Tools.Office.targets is unavailable on this machine.
```

Incorrect:

```text
PR claims ETABS/SAP2000 smoke tests passed without running those applications.
```

## When to update this document

Update this checklist when CI changes, project split changes, or manual release/smoke requirements change.

## Related documents

- [git-convention.md](git-convention.md)
- [testing-convention.md](testing-convention.md)
- [release-convention.md](release-convention.md)

