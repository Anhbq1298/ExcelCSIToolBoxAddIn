# Release Convention

Releases must account for .NET Framework 4.8, VSTO packaging, signing, Excel startup behavior, and ETABS/SAP2000 integration smoke tests.

## Build types

- Development build: local build for feature validation.
- Internal test build: signed or test-signed package for internal Excel/CSI smoke testing.
- Signed release build: production-signed VSTO package and release artifacts.

## Mandatory release steps

1. Confirm version number and release notes.
2. Restore and build all projects in the release environment.
3. Build the VSTO AddIn with Office/VSTO targets installed.
4. Sign manifests and assemblies according to the current certificate policy.
5. Never commit private certificate passwords or local signing secrets.
6. Generate installer or publish artifacts.
7. Smoke test Excel startup and AddIn load behavior.
8. Smoke test ETABS attach/read/write features affected by the release.
9. Smoke test SAP2000 attach/read/write features affected by the release.
10. Smoke test Excel import/export workflows.
11. Create GitHub Release artifacts and notes.
12. Document rollback steps.

## Release notes

Include user-facing changes, bug fixes, migration notes, known limitations, and manual smoke coverage. Do not claim manual ETABS, SAP2000, Excel, or installer verification unless performed.

Correct:

```text
Known limitation: SAP2000 smoke test not run because SAP2000 was unavailable on the release machine.
```

Incorrect:

```text
All CSI products verified.
```

## When to update this document

Update this document when VSTO packaging, signing, certificate storage, installer generation, or release channels change.

## Related documents

- [git-convention.md](git-convention.md)
- [pull-request-checklist.md](pull-request-checklist.md)
- [logging-convention.md](logging-convention.md)

Checklist:

- Release environment has Office/VSTO targets.
- Signing is secure.
- Smoke tests are reported honestly.
- Rollback is documented.

