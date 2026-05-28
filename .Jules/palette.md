## 2026-05-24 - Added Tooltips and Accessibility Tags to WPF Elements
**Learning:** Adding screen reader support and general tooltips in WPF significantly improves accessibility and UX for forms.
**Action:** Always check interactive XAML elements for `AutomationProperties.Name`, `AutomationProperties.LabeledBy`, and `ToolTip` attributes.

## 2026-05-25 - Combining WPF DataTriggers and Grid Layouts for Empty States
**Learning:** WPF elements like `materialDesign:Card` can only have a single child. To overlay an empty state over a list component like `DataGrid`, you must wrap them in a `Grid`. Empty states provide better UX by informing the user instead of leaving a blank space. Using a `DataTrigger` bound to `Items.Count == 0` allows the empty state text to seamlessly display only when the collection is empty without needing code-behind logic.
**Action:** Always consider empty states for lists/tables, remembering to wrap siblings in layout containers (`Grid`) inside single-child parent nodes, and utilize XAML `DataTrigger` for visibility.
