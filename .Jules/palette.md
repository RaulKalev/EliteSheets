## 2026-05-24 - Added Tooltips and Accessibility Tags to WPF Elements
**Learning:** Adding screen reader support and general tooltips in WPF significantly improves accessibility and UX for forms.
**Action:** Always check interactive XAML elements for `AutomationProperties.Name`, `AutomationProperties.LabeledBy`, and `ToolTip` attributes.

## 2026-05-25 - Added Empty State for Filterable Lists
**Learning:** Providing explicit "empty state" feedback for lists and data grids significantly improves UX when searching or filtering yields no results, preventing users from wondering if the application is broken.
**Action:** Always verify if a list or grid component needs an empty state, and use UI triggers (e.g., `DataTrigger` on `Items.Count`) to display helpful messages like "No results found" when appropriate.
