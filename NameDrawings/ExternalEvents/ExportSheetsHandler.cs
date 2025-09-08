using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using EliteSheets.Exports;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EliteSheets.ExternalEvents
{
    public class ExportSheetsHandler : IExternalEventHandler
    {
        public UIDocument UiDoc { get; set; }
        public Document Doc { get; set; }
        public List<ViewSheet> SheetsToExport { get; set; } = new List<ViewSheet>();

        public string ExportPath { get; set; }
        public string ExportSetupName { get; set; }
        public bool ExportPdf { get; set; } = true;
        public bool ExportDwg { get; set; } = true;

        public void Execute(UIApplication app)
        {
            if (Doc == null || UiDoc == null || SheetsToExport == null || string.IsNullOrWhiteSpace(ExportPath))
                return;

            bool anySuccess = false;

            if (ExportDwg)
                anySuccess |= ExportAllDwgSheets();

            if (ExportPdf)
                anySuccess |= ExportAllPdfSheets();

            ShowCompletionDialog(anySuccess);
        }

        private bool ExportAllDwgSheets()
        {
            bool success = false;
            var options = DWGExportOptions.GetPredefinedOptions(Doc, ExportSetupName);
            var dwgExporter = new DwgExportService(Doc, options, ExportPath);

            foreach (var sheet in SheetsToExport)
            {
                try
                {
                    if (dwgExporter.ExportSheet(sheet))
                        success = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DWG export failed for {sheet.Name}: {ex.Message}");
                }
            }

            return success;
        }
        private bool ExportAllPdfSheets()
        {
            bool success = false;
            var pdfExporter = new PdfExportService(Doc);

            // Group by group number between "-7-" and "_"
            // And capture merge order from the number after the last "--"
            var grouped = new Dictionary<string, List<(ViewSheet Sheet, int Order)>>(StringComparer.OrdinalIgnoreCase);
            var singles = new List<ViewSheet>();

            foreach (var sheet in SheetsToExport)
            {
                string num = sheet.SheetNumber ?? string.Empty;

                // Only merge when there is an explicit order marker ("--N")
                if (!TryParseMergeOrder(num, out int order))
                {
                    singles.Add(sheet);
                    continue;
                }

                if (!TryParseGroupNumber(num, out string groupKey))
                {
                    // If we can’t find a group number, treat as single
                    singles.Add(sheet);
                    continue;
                }

                if (!grouped.TryGetValue(groupKey, out var list))
                {
                    list = new List<(ViewSheet, int)>();
                    grouped[groupKey] = list;
                }
                list.Add((sheet, order));
            }

            // Export singles (no "--" or failed parsing)
            foreach (var sheet in singles)
            {
                try
                {
                    if (pdfExporter.ExportSheetAsPdf(sheet, ExportPath))
                        success = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"PDF export failed for {sheet.Name}: {ex.Message}");
                }
            }

            // Export merged groups, sorted by the order number after "--"
            foreach (var kvp in grouped)
            {
                string groupNumber = kvp.Key; // e.g., "05"
                var orderedSheets = kvp.Value
                    .OrderBy(t => t.Order)
                    .ThenBy(t => t.Sheet.SheetNumber, StringComparer.OrdinalIgnoreCase)
                    .Select(t => t.Sheet)
                    .ToList();

                // Build a nice combined file name like "<prefix>-7-<group>.pdf"
                string outputName = BuildCombinedFileName(orderedSheets.First().SheetNumber, groupNumber);

                try
                {
                    if (pdfExporter.ExportCombinedPdf(orderedSheets, ExportPath, outputName))
                        success = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Combined PDF export failed for group '{groupNumber}': {ex.Message}");
                }
            }

            return success;
        }
        // e.g. "...--1" or "...-- 1" at the END
        private static bool TryParseMergeOrder(string sheetNumber, out int order)
        {
            order = int.MaxValue;
            if (string.IsNullOrWhiteSpace(sheetNumber)) return false;

            var m = Regex.Match(sheetNumber, @"--\s*(\d+)\s*$");
            if (!m.Success) return false;

            return int.TryParse(m.Groups[1].Value, out order);
        }

        // e.g. "...-7-05_..." -> group "05"
        // tolerant to spaces and different dash glyphs
        private static bool TryParseGroupNumber(string sheetNumber, out string group)
        {
            group = null;
            if (string.IsNullOrWhiteSpace(sheetNumber)) return false;

            // normalize some common dash variants to simple '-'
            var normalized = sheetNumber.Replace('–', '-').Replace('—', '-');

            var m = Regex.Match(normalized, @"-7-\s*([0-9]+)\s*_", RegexOptions.CultureInvariant);
            if (!m.Success) return false;

            group = m.Groups[1].Value.Trim();
            return group.Length > 0;
        }

        // "<prefix>-7-<group>_<title>" where:
        // prefix = everything BEFORE "-7-"
        // group  = digits after "-7-"
        // title  = text after the first "_" following the group, without the trailing "--N"
        private static string BuildCombinedFileName(string sheetNumber, string fallbackGroupNumber)
        {
            if (string.IsNullOrWhiteSpace(sheetNumber))
                return $"Group-{fallbackGroupNumber}";

            string prefix, group, title;
            if (TryParseParts(sheetNumber, out prefix, out group, out title))
                return $"{prefix}-7-{group}_{title}";

            // Fallback to previous behavior if parsing fails
            var normalized = sheetNumber.Replace('–', '-').Replace('—', '-');
            int p7 = normalized.IndexOf("-7-", StringComparison.Ordinal);
            if (p7 > 0)
            {
                string cleanPrefix = normalized.Substring(0, p7).TrimEnd('_', '-', ' ');
                return $"{cleanPrefix}-7-{fallbackGroupNumber}";
            }
            return $"Group-{fallbackGroupNumber}";
        }

        // Parses: "<prefix>-7-<group>_<title>[--N]" (dashes can be en/em)
        // Returns false if the pattern doesn’t match.
        private static bool TryParseParts(string sheetNumber, out string prefix, out string group, out string title)
        {
            prefix = group = title = null;
            if (string.IsNullOrWhiteSpace(sheetNumber)) return false;

            var normalized = sheetNumber.Replace('–', '-').Replace('—', '-');

            // ^(prefix)-7-(group)_(title)(--N)?$
            var m = Regex.Match(
                normalized,
                @"^(?<prefix>.+?)-7-\s*(?<group>\d+)\s*_(?<title>.+?)(?:--\s*\d+\s*)?$",
                RegexOptions.CultureInvariant);

            if (!m.Success) return false;

            prefix = m.Groups["prefix"].Value.TrimEnd('_', '-', ' ');
            group = m.Groups["group"].Value.Trim();
            title = m.Groups["title"].Value.Trim();

            // Defensive: avoid illegal filename chars if they ever appear in the title
            foreach (var c in Path.GetInvalidFileNameChars())
                title = title.Replace(c.ToString(), "");

            return prefix.Length > 0 && group.Length > 0 && title.Length > 0;
        }

        private void ShowCompletionDialog(bool anySuccess)
        {
            if (anySuccess)
            {
                TaskDialogResult result = TaskDialog.Show(
                    "Export lõppenud.",
                    "Export lõppenud.\n\nAvada ekspordi kaust?",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    TaskDialogResult.No);

                if (result == TaskDialogResult.Yes && Directory.Exists(ExportPath))
                {
                    Process.Start("explorer.exe", ExportPath);
                }
            }
            else
            {
                TaskDialog.Show("EliteSheets - Export Failed", "Export failed for all selected sheets.");
            }
        }

        public string GetName() => "Export Sheets Handler";
    }
}
