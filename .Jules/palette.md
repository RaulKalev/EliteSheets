## 2026-05-24 - Added Tooltips and Accessibility Tags to WPF Elements
**Learning:** Adding screen reader support and general tooltips in WPF significantly improves accessibility and UX for forms.
**Action:** Always check interactive XAML elements for `AutomationProperties.Name`, `AutomationProperties.LabeledBy`, and `ToolTip` attributes.
## 2025-05-24 - Empty States and Accessibility in WPF DataGrids
**Learning:** Adding empty states significantly improves the user experience when a search yields no results. WPF's DataTrigger allows adding these easily without code-behind. Screen reader support is vital.
**Action:** Use DataTriggers on `HasItems` property to toggle the visibility of empty state text blocks overlaid on DataGrids. Continue to ensure `AutomationProperties.Name` is added for accessibility.
