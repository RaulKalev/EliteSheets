## 2026-05-24 - Added Tooltips and Accessibility Tags to WPF Elements
**Learning:** Adding screen reader support and general tooltips in WPF significantly improves accessibility and UX for forms.
**Action:** Always check interactive XAML elements for `AutomationProperties.Name`, `AutomationProperties.LabeledBy`, and `ToolTip` attributes.
## 2024-05-17 - MaterialDesignThemes Card Empty States
**Learning:** `materialDesign:Card` in WPF only accepts a single child element. Attempting to add a sibling element directly inside it for an empty state causes an MC3089 build error.
**Action:** When adding empty states to DataGrids within a Card, wrap the DataGrid and the empty state element in a layout container like `Grid`. Use a `DataTrigger` bound to `Items.Count == 0` on the DataGrid to conditionally show the empty state element.
