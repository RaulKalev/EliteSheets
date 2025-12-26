using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EliteSheets.Services
{
    public class SheetGroupingService
    {
        public class PartitionResult
        {
            /// <summary>
            /// Sheets that are not part of any merged group (or are single files).
            /// </summary>
            public List<ViewSheet> Singles { get; } = new List<ViewSheet>();

            /// <summary>
            /// Sheets grouped by their group ID (e.g. "05").
            /// Value contains (Sheet, MergeOrder).
            /// </summary>
            public Dictionary<string, List<(ViewSheet Sheet, int Order)>> Groups { get; } 
                = new Dictionary<string, List<(ViewSheet, int)>>(StringComparer.OrdinalIgnoreCase);
        }

        public PartitionResult Partition(IEnumerable<ViewSheet> sheets)
        {
            var result = new PartitionResult();

            foreach (var sheet in sheets)
            {
                string num = sheet.SheetNumber ?? string.Empty;

                // Test for merge order "--N"
                if (!TryParseMergeOrder(num, out int order))
                {
                    result.Singles.Add(sheet);
                    continue;
                }

                // Test for group ID "-7-XX_"
                if (!TryParseGroupNumber(num, out string groupKey))
                {
                    result.Singles.Add(sheet);
                    continue;
                }

                if (!result.Groups.TryGetValue(groupKey, out var list))
                {
                    list = new List<(ViewSheet, int)>();
                    result.Groups[groupKey] = list;
                }
                list.Add((sheet, order));
            }

            return result;
        }

        // e.g. "...--1" or "...-- 1" at the END
        public bool TryParseMergeOrder(string sheetNumber, out int order)
        {
            order = int.MaxValue;
            if (string.IsNullOrWhiteSpace(sheetNumber)) return false;

            var m = Regex.Match(sheetNumber, @"--\s*(\d+)\s*$");
            if (!m.Success) return false;

            return int.TryParse(m.Groups[1].Value, out order);
        }

        // e.g. "...-7-05_..." -> group "05"
        public bool TryParseGroupNumber(string sheetNumber, out string group)
        {
            group = null;
            if (string.IsNullOrWhiteSpace(sheetNumber)) return false;

            var normalized = sheetNumber.Replace('–', '-').Replace('—', '-');
            var m = Regex.Match(normalized, @"-7-\s*([0-9]+)\s*_", RegexOptions.CultureInvariant);
            if (!m.Success) return false;

            group = m.Groups[1].Value.Trim();
            return group.Length > 0;
        }

        public string BuildCombinedFileName(string sheetNumber, string fallbackGroupNumber)
        {
            if (string.IsNullOrWhiteSpace(sheetNumber))
                return $"Group-{fallbackGroupNumber}";

            string prefix, group, title;
            if (TryParseParts(sheetNumber, out prefix, out group, out title))
                return $"{prefix}-7-{group}_{title}";

            // Fallback
            var normalized = sheetNumber.Replace('–', '-').Replace('—', '-');
            int p7 = normalized.IndexOf("-7-", StringComparison.Ordinal);
            if (p7 > 0)
            {
                string cleanPrefix = normalized.Substring(0, p7).TrimEnd('_', '-', ' ');
                return $"{cleanPrefix}-7-{fallbackGroupNumber}";
            }
            return $"Group-{fallbackGroupNumber}";
        }

        private bool TryParseParts(string sheetNumber, out string prefix, out string group, out string title)
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

            foreach (var c in Path.GetInvalidFileNameChars())
                title = title.Replace(c.ToString(), "");

            return prefix.Length > 0 && group.Length > 0 && title.Length > 0;
        }
    }
}
