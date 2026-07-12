# Git Convention

Use small, reviewable commits. Separate structural moves from behavior changes whenever possible.

## Commit style

Use Conventional Commit style where practical:

```text
refactor(core): move CSI contracts out of Data project
refactor(ui): group analysis result components by feature
refactor(csi): split ETABS services by capability
docs(skills): add Excel interop development convention
test(architecture): enforce project dependency boundaries
```

## Mandatory rules

- Do not combine the entire migration or feature into one commit.
- Keep refactor, behavior change, tests, and docs separable unless they are inseparable.
- Do not include unrelated formatting churn.
- Do not commit secrets, PFX passwords, local user paths, or generated build outputs.
- Treat interop DLLs in `lib/` as deliberate binary dependencies; do not replace them casually.
- Generated files must be identified by name or folder, such as `CsiMethodCatalog.generated.cs`.
- Commit messages must say what changed and where.

Branch names should be short and scoped:

```text
feature/analysis-export
fix/excel-range-cancel
refactor/csi-folders
docs/skills-conventions
```

## Migration commit order

For future structural work, commit in this order: baseline docs, solution/project moves, ownership moves, feature folders, UI folders, tests, architecture checks, docs.

Correct:

```text
refactor(application): organize connectivity use cases by feature
```

Incorrect:

```text
update stuff
```

## When to update this document

Update this document when release branching, signing artifacts, generated-file policy, or PR workflow changes.

## Related documents

- [pull-request-checklist.md](pull-request-checklist.md)
- [release-convention.md](release-convention.md)
- [coding-agent-instructions.md](coding-agent-instructions.md)

Checklist:

- Commit is scoped.
- Tests or docs accompany risky changes.
- No unrelated files are included.
- Generated/binary changes are intentional.

