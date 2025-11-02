using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using EliteSheets.Exports;
using EliteSheets.Services;
using netDxf;
using netDxf.Blocks;
using netDxf.Entities;
using netDxf.Objects;
using netDxf.Units;
using Newtonsoft.Json;
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
        public bool ExportDxf { get; set; } = false;
        public string TemplateDxfPath { get; set; }

        public void Execute(UIApplication app)
        {
            if (Doc == null || UiDoc == null || SheetsToExport == null || string.IsNullOrWhiteSpace(ExportPath))
                return;

            bool anySuccess = false;

            if (ExportDwg)
                anySuccess |= ExportAllDwgSheets();

            if (ExportPdf)
                anySuccess |= ExportAllPdfSheets();

            if (ExportDxf)
                anySuccess |= ExportAllDxfSheets();

            ShowCompletionDialog(anySuccess);
        }
        private static string LoadTemplatePathFromConfig()
        {
            const string configFile = @"C:\ProgramData\RK Tools\EliteSheets\config.json";
            try
            {
                if (!File.Exists(configFile)) return null;
                var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(File.ReadAllText(configFile));
                if (dict != null && dict.TryGetValue("TemplateDxfPath", out var v))
                {
                    var p = v?.ToString();
                    return !string.IsNullOrWhiteSpace(p) && File.Exists(p) ? p : null;
                }
            }
            catch { }
            return null;
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
        private bool ExportAllDxfSheets()
        {
            bool anyExportSuccess = false;
            var postErrors = new List<string>();

            // --- Partition into singles vs grouped like PDF ---
            var grouped = new Dictionary<string, List<(ViewSheet Sheet, int Order)>>(StringComparer.OrdinalIgnoreCase);
            var singles = new List<ViewSheet>();

            foreach (var sheet in SheetsToExport)
            {
                var num = sheet.SheetNumber ?? string.Empty;

                if (!TryParseMergeOrder(num, out int order))
                {
                    singles.Add(sheet); // no "--N" -> treat as single
                    continue;
                }

                if (!TryParseGroupNumber(num, out string groupKey))
                {
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

            // --- Helper: find DXF by sheet number in a directory ---
            string FindDxfForSheet(ViewSheet sheet, string folder)
            {
                var num = sheet.SheetNumber ?? "";

                // Typical Revit pattern: "Prefix-Sheet - <SheetNumber> - <SheetName>.dxf"
                foreach (var fp in Directory.EnumerateFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly))
                {
                    var fn = Path.GetFileNameWithoutExtension(fp);
                    if (fn.IndexOf($" - {num} - ", StringComparison.OrdinalIgnoreCase) >= 0)
                        return fp;
                }

                // Fallback: filename contains the sheet number somewhere
                foreach (var fp in Directory.EnumerateFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly))
                {
                    var fn = Path.GetFileNameWithoutExtension(fp);
                    if (fn.IndexOf(num, StringComparison.OrdinalIgnoreCase) >= 0)
                        return fp;
                }
                return null;
            }

            try
            {
                var dxfExporter = new EliteSheets.Services.DxfExportService();
                var promoter = new EliteSheets.Services.DxfPaperToModelPromoter();
                var merger = new EliteSheets.Services.DxfMergeService();

                // Resolve template path: prefer value passed by MainWindow, else read JSON.
                string templatePath = (this is EliteSheets.ExternalEvents.ExportSheetsHandler h && !string.IsNullOrWhiteSpace(h.TemplateDxfPath))
                    ? h.TemplateDxfPath
                    : LoadTemplatePathFromConfig(); // <-- include the helper from earlier answer

                if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                {
                    TaskDialog.Show("DXF eksport",
                        "Mallifaili (DXF) asukoht ei ole seadistatud või faili ei leitud. " +
                        "Ava EliteSheets aken ja vali mall seadetes.");
                    return anyExportSuccess;
                }

                // ============ 1) Export SINGLES directly to the final export folder ============
                var singleIds = singles.Select(s => s.Id).ToList();
                if (singleIds.Count > 0)
                {
                    // Snapshot to detect new singles quickly
                    var pre = new HashSet<string>(
                        Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly),
                        StringComparer.OrdinalIgnoreCase);

                    if (!dxfExporter.Export(Doc, singleIds, ExportPath, "DXF_Sheets", ExportSetupName, false, out string failureSingles))
                    {
                        Debug.WriteLine($"DXF export (singles) failed: {failureSingles}");
                    }
                    else
                    {
                        anyExportSuccess = true;

                        // Promote Paper→Model for the new single files only
                        var newSingles = Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly)
                                                  .Where(p => !pre.Contains(p))
                                                  .ToList();

                        // If snapshot didn’t find them (timestamp fallback)
                        if (newSingles.Count == 0)
                        {
                            var cutoff = DateTime.UtcNow.AddMinutes(-3);
                            newSingles = Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly)
                                                  .Where(p => File.GetLastWriteTimeUtc(p) >= cutoff)
                                                  .ToList();
                        }

                        foreach (var path in newSingles)
                        {
                            try { promoter.PromotePaperToModel(path); }
                            catch (Exception ex) { postErrors.Add($"{Path.GetFileName(path)}: {ex.Message}"); }
                        }
                    }
                }

                // ============ 2) Export GROUPED sheets to a TEMP folder, merge, then delete ============
                if (grouped.Count > 0)
                {
                    // Separate temp folder for grouped sheets so we can safely delete them afterwards
                    string tempRoot = Path.Combine(ExportPath, "_tmp_dxf_merge_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(tempRoot);

                    try
                    {
                        // Collect all sheet Ids that belong to any group
                        var groupedIds = grouped.Values.SelectMany(v => v.Select(t => t.Sheet.Id)).Distinct().ToList();

                        if (groupedIds.Count > 0)
                        {
                            if (!dxfExporter.Export(Doc, groupedIds, tempRoot, "DXF_Sheets", ExportSetupName, false, out string failureGroups))
                            {
                                Debug.WriteLine($"DXF export (groups) failed: {failureGroups}");
                            }
                            else
                            {
                                anyExportSuccess = true;

                                // Promote all temp DXFs
                                foreach (var path in Directory.EnumerateFiles(tempRoot, "*.dxf", SearchOption.TopDirectoryOnly))
                                {
                                    try { promoter.PromotePaperToModel(path); }
                                    catch (Exception ex) { postErrors.Add($"(temp) {Path.GetFileName(path)}: {ex.Message}"); }
                                }

                                // Merge per group using ONLY temp files → output to ExportPath
                                foreach (var kvp in grouped)
                                {
                                    string groupNumber = kvp.Key; // e.g. "01"
                                    var orderedSheets = kvp.Value
                                        .OrderBy(t => t.Order)
                                        .ThenBy(t => t.Sheet.SheetNumber, StringComparer.OrdinalIgnoreCase)
                                        .Select(t => t.Sheet)
                                        .ToList();

                                    var paths = new List<string>();
                                    foreach (var s in orderedSheets)
                                    {
                                        var p = FindDxfForSheet(s, tempRoot);
                                        if (!string.IsNullOrEmpty(p) && File.Exists(p))
                                            paths.Add(p);
                                        else
                                            postErrors.Add($"DXF for sheet '{s.SheetNumber}' not found for merging (temp).");
                                    }
                                    if (paths.Count == 0) continue;

                                    string combinedNameNoExt = BuildCombinedFileName(orderedSheets.First().SheetNumber, groupNumber);
                                    string combinedPath = Path.Combine(ExportPath, combinedNameNoExt + ".dxf");

                                    try
                                    {
                                        // Merge sources → into TEMPLATE model space → save as combinedPath
                                        merger.MergeIntoTemplate(
                                            paths,
                                            templatePath,
                                            combinedPath,
                                            sheetSpacingMm: 220.0,
                                            insertXmm: 0.0,
                                            insertYmm: 0.0
                                        );
                                    }
                                    catch (Exception ex)
                                    {
                                        postErrors.Add($"Merge into template '{combinedNameNoExt}.dxf' failed: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Delete ALL temp files/folder (so only the merged DXFs remain)
                        try { Directory.Delete(tempRoot, true); }
                        catch (Exception ex) { postErrors.Add($"Temp cleanup failed: {ex.Message}"); }
                    }
                }
            }
            catch (Exception ex)
            {
                postErrors.Add($"DXF export/post exception: {ex.Message}");
            }

            if (postErrors.Count > 0)
            {
                TaskDialog.Show("DXF post-processing",
                    "Some DXF files were exported/merged but issues occurred:\n\n" + string.Join("\n", postErrors));
            }

            return anyExportSuccess;
        }



        /// <summary>
        /// Merge multiple DXFs into one DXF by inserting each file's Model Space as a block,
        /// spaced apart along +X (so sheets don't overlap). Units: millimetres.
        /// </summary>
        private static void MergeDxfFilesSideBySide(IList<string> sourcePaths, string outputPath)
        {
            if (sourcePaths == null || sourcePaths.Count == 0)
                throw new ArgumentException("No DXF files to merge.");

            const double PAGE_SPACING_MM = 100000.0; // horizontal spacing between merged “pages”

            var target = new DxfDocument();
            target.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            double offsetX = 0.0;
            int idx = 1;

            foreach (var path in sourcePaths)
            {
                var src = DxfDocument.Load(path);

                // find source MODEL space block
                Block srcModel =
                    src.Blocks.Contains(Layout.ModelSpaceName)
                        ? src.Blocks[Layout.ModelSpaceName]
                        : src.Blocks.FirstOrDefault(b =>
                              string.Equals(b.Name, "*Model_Space", StringComparison.OrdinalIgnoreCase));

                if (srcModel == null)
                    throw new InvalidDataException($"Model Space block not found in: {Path.GetFileName(path)}");

                // make a new block in the TARGET to host cloned entities from this source's MODEL space
                string blockName = $"SHEET_{idx}_{Guid.NewGuid():N}";
                var blk = new Block(blockName);

                foreach (var ent in srcModel.Entities)
                    blk.Entities.Add((EntityObject)ent.Clone());

                // register block + insert it at offset
                target.Blocks.Add(blk);
                target.Entities.Add(new Insert(blk)
                {
                    Position = new netDxf.Vector3(offsetX, 0, 0),
                    Scale = new netDxf.Vector3(1, 1, 1),
                    Rotation = 0
                });

                offsetX += PAGE_SPACING_MM;
                idx++;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
            target.Save(outputPath);
        }

        private static string Sanitize(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "NA";
            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                s = s.Replace(c.ToString(), "");
            return s;
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
