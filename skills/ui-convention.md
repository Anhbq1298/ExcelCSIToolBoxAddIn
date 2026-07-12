# UI Convention

UI code currently lives in `src/ExcelCSIToolBoxAddIn/UI`. It is grouped by user-facing module:

```text
UI/Modules/AiAgent
UI/Modules/AnalysisResults
UI/Modules/Loadings
UI/Modules/ModellingHelpers
UI/Modules/Sections
UI/Modules/ToolboxShell
UI/Shared
```

## Mandatory rules

- Preserve the existing WPF/WinForms visual language unless the task explicitly changes UI design.
- Keep related views, ViewModels, dialogs, controls, profiles, and local UI models close to the feature module.
- Use `UI/Shared` only for reusable UI capabilities such as commands, Excel selection UI, host controls, progress, themes, and base ViewModels.
- Read-only operations should not ask for confirmation unless they can be expensive or confusing.
- Model write operations need clear confirmation or preview behavior.
- Progress UI should be used for batch or long-running work.
- Empty, loading, error, disabled, and in-development states must be visible and actionable.
- Windows should have consistent owner behavior through the AddIn host/window manager.
- Avoid one-off styles; reuse `UI/Shared/Themes/EtabsToolboxTheme.xaml` where possible.

## Layout patterns

Read-only export dialog:

```text
Header/status -> options -> preview/selection summary -> primary Export button -> Cancel
```

Excel import dialog:

```text
Range picker -> validation summary -> parsed grid preview -> Import/Cancel
```

Model modification dialog:

```text
Connection/model status -> inputs -> validation -> preview/dry run -> Apply/Cancel
```

Batch result dialog:

```text
Summary counts -> item grid -> failed item details -> Copy/Close
```

Feature under development panel:

```text
Short status -> disabled action -> no fake controls that appear executable
```

Correct:

```text
UI/Modules/Loadings/ShellUniformLoadSetForm.cs
UI/Shared/Progress/BatchProgressWindow.xaml
```

Incorrect:

```text
UI/Views/ShellUniformLoadSetForm.cs
UI/Misc/BatchProgressWindow.xaml
```

## When to update this document

Update this document when adding a new UI module, changing theme resources, or changing confirmation/progress behavior.

## Related documents

- [viewmodel-convention.md](viewmodel-convention.md)
- [excel-interop-convention.md](excel-interop-convention.md)
- [dependency-injection-convention.md](dependency-injection-convention.md)

Checklist:

- Feature UI is in `UI/Modules/<Module>`.
- Shared UI code is genuinely shared.
- Write actions are confirmed or previewed.
- Resource paths are updated after moves.

