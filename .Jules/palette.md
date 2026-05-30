## 2026-05-24 - Added Tooltips and Accessibility Tags to WPF Elements
**Learning:** Adding screen reader support and general tooltips in WPF significantly improves accessibility and UX for forms.
**Action:** Always check interactive XAML elements for `AutomationProperties.Name`, `AutomationProperties.LabeledBy`, and `ToolTip` attributes.

## 2026-05-30 - Adding Empty States in WPF DataGrid
**Learning:** When adding empty states overlayed onto a WPF DataGrid inside a MaterialDesign Card, you must wrap them in a Grid to avoid single-child limitations.
**Action:** Use a TextBlock inside a Grid overlay and control visibility using a `DataTrigger` bound to `Items.Count == 0` to ensure clear feedback when no records are present.
