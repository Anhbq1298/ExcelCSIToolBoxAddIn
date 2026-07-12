# Excel CSI Toolbox Skills

This folder is the development playbook for the Excel CSI Toolbox repository. It explains where code belongs, how dependencies flow, how ETABS/SAP2000 and Excel COM code are isolated, and how future coding agents should work safely in this codebase.

The top-level Markdown files are the current conventions. Existing subfolders such as `development/`, `etabs-api-integration/`, and `ui-management/` are retained as skill sources, but the documents listed here are the canonical repository conventions after the folder refactor.

## How to use these documents

Developers should read the relevant convention before adding or moving code. Coding agents must read [coding-agent-instructions.md](coding-agent-instructions.md) first, then follow the document priority below.

Document priority order:

1. [coding-agent-instructions.md](coding-agent-instructions.md)
2. [architecture-convention.md](architecture-convention.md)
3. [repository-structure.md](repository-structure.md)
4. Feature-specific conventions such as [csi-api-convention.md](csi-api-convention.md), [excel-interop-convention.md](excel-interop-convention.md), [ui-convention.md](ui-convention.md), and [viewmodel-convention.md](viewmodel-convention.md)
5. [coding-convention.md](coding-convention.md)

When documents conflict, use the higher-priority document and update the lower-priority document in the same change or call out the mismatch in the PR.

## Required reading before developing a feature

- Every task: [coding-agent-instructions.md](coding-agent-instructions.md), [architecture-convention.md](architecture-convention.md), and [repository-structure.md](repository-structure.md)
- Application workflow: [application-logic-convention.md](application-logic-convention.md)
- ETABS or SAP2000 work: [csi-api-convention.md](csi-api-convention.md)
- Excel read/write work: [excel-interop-convention.md](excel-interop-convention.md)
- WPF or WinForms work: [ui-convention.md](ui-convention.md) and [viewmodel-convention.md](viewmodel-convention.md)
- Tests: [testing-convention.md](testing-convention.md)
- Commits and PRs: [git-convention.md](git-convention.md) and [pull-request-checklist.md](pull-request-checklist.md)

## Document index

- [repository-structure.md](repository-structure.md)
- [architecture-convention.md](architecture-convention.md)
- [coding-convention.md](coding-convention.md)
- [application-logic-convention.md](application-logic-convention.md)
- [csi-api-convention.md](csi-api-convention.md)
- [excel-interop-convention.md](excel-interop-convention.md)
- [ui-convention.md](ui-convention.md)
- [viewmodel-convention.md](viewmodel-convention.md)
- [feature-development-guide.md](feature-development-guide.md)
- [testing-convention.md](testing-convention.md)
- [error-handling-convention.md](error-handling-convention.md)
- [logging-convention.md](logging-convention.md)
- [naming-convention.md](naming-convention.md)
- [dependency-injection-convention.md](dependency-injection-convention.md)
- [git-convention.md](git-convention.md)
- [pull-request-checklist.md](pull-request-checklist.md)
- [release-convention.md](release-convention.md)
- [coding-agent-instructions.md](coding-agent-instructions.md)

## Updating conventions

Update these documents when a project boundary changes, a new integration pattern is introduced, a repeated review comment becomes a rule, or architecture tests are added or changed. Prefer changing the most specific document and linking to it rather than duplicating long guidance.

Correct:

```text
Update csi-api-convention.md when a new ETABS return-code pattern is introduced, then link to it from the PR.
```

Incorrect:

```text
Add an undocumented convention in a PR comment and leave the skills folder stale.
```

## Related documents

- [coding-agent-instructions.md](coding-agent-instructions.md)
- [architecture-convention.md](architecture-convention.md)
- [repository-structure.md](repository-structure.md)

Checklist:

- The rule matches the current code.
- Correct and incorrect examples use real repository paths.
- Related documents are linked.
- Architecture tests are updated when the rule is enforceable.
