---
name: ui-management
description: Manage WPF UI and MVVM changes for the ExcelCSIToolBoxAddIn ETABS/SAP2000 toolbox. Use when adjusting layout, styles, bindings, navigation, tool windows, buttons, tree views, or UI behavior while preserving backend workflows.
---

# UI Management

## Visual Style

Use the repo's existing WPF theme resources first, especially `UI/Themes/EtabsToolboxTheme.xaml`.

Preferred feel:

- clean engineering software
- Microsoft Office / Windows desktop aligned
- navy-white palette
- readable and compact
- professional rather than decorative

Use `Segoe UI` as the UI font.

## Palette

Use these colors when adding or updating shared resources:

- Main navy: `#0B1F3A`
- Header navy: `#102A4C`
- Sidebar navy: `#132B47`
- Card background: `#FFFFFF`
- Secondary card background: `#F7FAFC`
- Border: `#D6E0EA`
- Selected light blue: `#D8ECFF`
- Hover light blue: `#EEF7FF`
- Active blue border: `#2F80ED`
- Primary text: `#1F2937`
- Secondary text: `#64748B`
- Text on navy: `#FFFFFF`
- Muted text on navy: `#D7E3F0`

## Binding Safety

- Do not rename commands, properties, or DataContext assumptions unless updating every binding.
- Prefer existing commands and ViewModel patterns.
- Keep tree navigation separate from action buttons.
- Do not trigger ETABS extraction from tree selection unless explicitly requested.
- Use `CommandParameter="{Binding}"` when buttons represent model items.
- Keep `RelativeSource` command binding patterns consistent in `ItemsControl` templates.

## Toolbox Navigation Areas

Preserve these areas in the tree and breadcrumbs:

- `ANALYSIS RESULTS`: result/table buttons such as Joint Output, Element Output, Structure Output.
- `MISCELLANEOUS DATA`: project information and material list exports. Keep this category visibly separate from Analysis Results.
- `Modelling Helper`: helper workflows such as array creation and shell creation from selected frames.
- `Element Manipulation`: point, frame, shell, and object connectivity tools.
- `Model`: general information, section property, load pattern, load combination, and stiffness modifier pages.

When updating the shared content area, make sure `ActiveTableCategory`, `ActiveAnalysisResultsGroup`, `ActivePageTitle`, and `ActivePageBreadcrumb` still communicate the selected area correctly. Do not let Miscellaneous Data appear as an Analysis Results breadcrumb unless intentionally requested.

## Layout Rules

- Preserve current layout structure unless the task explicitly asks for redesign.
- Prefer shared styles in resource dictionaries over repeated inline styling.
- Keep text readable on all selected and hover states.
- Selected navigation/button state should be light blue with dark navy text.
- Never use white or gray text on selected light-blue backgrounds.
- Avoid changing backend extraction, DTOs, service logic, or project references while doing visual-only UI work.

## WPF Checks

After UI changes:

- Build the solution to validate XAML compile.
- Scan for broken binding names if commands or ViewModel properties were touched.
- Check that Attach, Close, Lock Model, Unit System, and Analysis Results navigation still bind to existing commands.
- For Analysis Results, verify button content binds to `Title` and button command receives the full `AnalysisResultItem`.
