# UI Skills - ExcelToolbox CSI Addin

## General Style

Use a clean, professional navy-white UI style suitable for engineering software tools.

The UI should feel:

* clean
* technical
* modern
* readable
* consistent with Microsoft Office / Windows desktop tooling

Do not use heavy black-white contrast unless required by the host application.

## Font

Use one font only:

Segoe UI

Do not use fallback font chains.

Apply the font at the root window or application resource level when possible.

## Color Palette

Main navy background: #0B1F3A
Header navy: #102A4C
Sidebar navy: #132B47
Card background: #FFFFFF
Secondary card background: #F7FAFC
Border color: #D6E0EA
Selected light blue: #D8ECFF
Hover light blue: #EEF7FF
Active blue border: #2F80ED
Primary text on white: #1F2937
Secondary text: #64748B
Text on navy: #FFFFFF
Muted text on navy: #D7E3F0

## Buttons

Normal:

* Background: #FFFFFF
* Text: #1F2937
* Border: #D6E0EA

Hover:

* Background: #EEF7FF
* Border: #9CCBFF
* Text: #0B1F3A

Selected:

* Background: #D8ECFF
* Border: #2F80ED
* Text: #0B1F3A
* FontWeight: SemiBold

Disabled:

* Background: #E5E7EB
* Text: #9CA3AF
* Border: #D1D5DB

## Selected State Rule

Whenever a button, menu item, tree item, or navigation item is selected:

* Background must become light blue.
* Text must remain dark navy.
* Text must have strong contrast.
* Never use white text on light blue.
* Never use gray text on light blue.

## Navigation

Navigation hierarchy must remain clear and readable.

Selected navigation items should use:

* Background: #D8ECFF
* Text: #0B1F3A
* Border/accent: #2F80ED

Do not hide or reduce visibility of tree arrows, icons, or item labels.

## Contrast Rules

Use dark navy or primary text on light surfaces. Use white or muted navy text on navy surfaces.

Do not use low-contrast gray text on selected or hover backgrounds.

## Layout Safety Rule

When improving UI style:

* Do not change existing layout structure unless explicitly requested.
* Do not rename existing commands.
* Do not break MVVM bindings.
* Do not modify backend extraction logic.
* Do not modify DTOs, database models, or service logic unless strictly required.
* If DTOs or shared models are changed, update all related mappings/usages accordingly.

## WPF Implementation Notes

Prefer shared styles in ResourceDictionary or Window.Resources.

Avoid repeating hardcoded colors across many controls.

Use reusable styles for:

* Window background
* Header
* Sidebar
* Content cards
* Buttons
* Navigation items
* TextBlocks
