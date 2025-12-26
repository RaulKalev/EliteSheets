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

        private readonly SheetGroupingService _groupingService = new SheetGroupingService();

        public void Execute(UIApplication app)
        {
            if (Doc == null || UiDoc == null || SheetsToExport == null || string.IsNullOrWhiteSpace(ExportPath))
                return;

            bool anySuccess = false;

            // 1. Partition sheets into Singles vs Groups
            var partition = _groupingService.Partition(SheetsToExport);
            
            // 2. DWG Export
            if (ExportDwg)
            {
                // Singles -> DWG
                if (partition.Singles.Count > 0)
                    anySuccess |= ExportDwgSingles(partition.Singles);

                // Groups -> DXF Merge (Smart Switching!)
                if (partition.Groups.Count > 0)
                    anySuccess |= ExportDxfGroups(partition.Groups);
            }

            // 3. DXF Export
            // Avoid re-running groups if already handled by Smart Switching above
            if (ExportDxf)
            {
                // Singles -> DXF
                if (partition.Singles.Count > 0)
                    anySuccess |= ExportDxfSingles(partition.Singles);

                // Groups -> DXF Merge (only if not already done by DWG logic)
                // If ExportDwg is true, we already exported groups above.
                if (!ExportDwg && partition.Groups.Count > 0)
                    anySuccess |= ExportDxfGroups(partition.Groups);
            }

            // 4. PDF Export
            if (ExportPdf)
            {
                anySuccess |= ExportAllPdfSheets(partition);
            }

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

        // --- DWG Logic ---

        private bool ExportDwgSingles(List<ViewSheet> singles)
        {
            bool success = false;
            var options = DWGExportOptions.GetPredefinedOptions(Doc, ExportSetupName);
            var dwgExporter = new DwgExportService(Doc, options, ExportPath);

            foreach (var sheet in singles)
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

        // --- PDF Logic ---

        private bool ExportAllPdfSheets(SheetGroupingService.PartitionResult partition)
        {
            bool success = false;
            var pdfExporter = new PdfExportService(Doc);

            // Singles
            foreach (var sheet in partition.Singles)
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

            // Groups
            foreach (var kvp in partition.Groups)
            {
                string groupNumber = kvp.Key;
                var orderedSheets = kvp.Value
                    .OrderBy(t => t.Order)
                    .ThenBy(t => t.Sheet.SheetNumber, StringComparer.OrdinalIgnoreCase)
                    .Select(t => t.Sheet)
                    .ToList();

                string outputName = _groupingService.BuildCombinedFileName(orderedSheets.First().SheetNumber, groupNumber);

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

        // --- DXF Logic ---

        /// <summary>
        /// Exports single sheets as individual DXF files.
        /// </summary>
        private bool ExportDxfSingles(List<ViewSheet> singles)
        {
            bool success = false;
            var dxfExporter = new EliteSheets.Services.DxfExportService();
            var promoter = new EliteSheets.Services.DxfPaperToModelPromoter();

            var ids = singles.Select(s => s.Id).ToList();
            if (ids.Count == 0) return false;

            // Snapshot existing DXFs to identify new ones
            var pre = new HashSet<string>(
                Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly),
                StringComparer.OrdinalIgnoreCase);

            if (!dxfExporter.Export(Doc, ids, ExportPath, "DXF_Sheets", ExportSetupName, false, out string failureMsg))
            {
                 Debug.WriteLine($"DXF export (singles) failed: {failureMsg}");
            }
            else
            {
                 success = true;
                 
                 // Identify newly created files
                 var newFiles = Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly)
                                         .Where(p => !pre.Contains(p))
                                         .ToList();
                 
                 if (newFiles.Count == 0)
                 {
                     var cutoff = DateTime.UtcNow.AddMinutes(-2);
                     newFiles = Directory.EnumerateFiles(ExportPath, "*.dxf", SearchOption.TopDirectoryOnly)
                                         .Where(p => File.GetLastWriteTimeUtc(p) >= cutoff)
                                         .ToList();
                 }

                 foreach(var f in newFiles)
                 {
                     try { promoter.PromotePaperToModel(f); }
                     catch(Exception ex) { Debug.WriteLine($"Promote failed for {f}: {ex.Message}"); }
                 }
            }

            return success;
        }

        /// <summary>
        /// Exports groups as MERGED DXF files (side-by-side).
        /// </summary>
        private bool ExportDxfGroups(Dictionary<string, List<(ViewSheet Sheet, int Order)>> groups)
        {
            bool success = false;
            var postErrors = new List<string>();

            // Resolve template path
            string templatePath = !string.IsNullOrWhiteSpace(TemplateDxfPath)
                ? TemplateDxfPath
                : LoadTemplatePathFromConfig();

            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                TaskDialog.Show("DXF eksport merge",
                    "Mallifaili (DXF) asukoht ei ole seadistatud või faili ei leitud. " +
                    "Ava EliteSheets aken ja vali mall seadetes.");
                return false;
            }

            // Temp folder
            string tempRoot = Path.Combine(ExportPath, "_tmp_dxf_merge_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);

            var dxfExporter = new EliteSheets.Services.DxfExportService();
            var promoter = new EliteSheets.Services.DxfPaperToModelPromoter();
            var merger = new EliteSheets.Services.DxfMergeService();

            try
            {
                var allIds = groups.Values.SelectMany(v => v.Select(t => t.Sheet.Id)).Distinct().ToList();
                if (allIds.Count > 0)
                {
                    if (dxfExporter.Export(Doc, allIds, tempRoot, "DXF_Sheets", ExportSetupName, false, out string failMsg))
                    {
                        success = true; // at least exported locally
                        
                        // Promote all in temp
                        foreach(var f in Directory.EnumerateFiles(tempRoot, "*.dxf"))
                        {
                            try { promoter.PromotePaperToModel(f); }
                            catch (Exception ex) { postErrors.Add($"(temp) {Path.GetFileName(f)}: {ex.Message}"); }
                        }

                        // Merge
                        foreach (var kvp in groups)
                        {
                            string groupNumber = kvp.Key;
                            var orderedSheets = kvp.Value
                                .OrderBy(t => t.Order)
                                .ThenBy(t => t.Sheet.SheetNumber, StringComparer.OrdinalIgnoreCase)
                                .Select(t => t.Sheet)
                                .ToList();

                            var sourcePaths = new List<string>();
                            foreach (var s in orderedSheets)
                            {
                                var p = FindDxfForSheet(s, tempRoot);
                                if (!string.IsNullOrEmpty(p) && File.Exists(p))
                                    sourcePaths.Add(p);
                                else
                                    postErrors.Add($"DXF for sheet '{s.SheetNumber}' not found for merging.");
                            }

                            if (sourcePaths.Count == 0) continue;

                            string combinedName = _groupingService.BuildCombinedFileName(orderedSheets.First().SheetNumber, groupNumber);
                            string outPath = Path.Combine(ExportPath, combinedName + ".dxf");

                            try
                            {
                                merger.MergeIntoTemplate(
                                    sourcePaths,
                                    templatePath,
                                    outPath,
                                    sheetSpacingMm: 220.0,
                                    insertXmm: 0.0,
                                    insertYmm: 0.0
                                );
                            
                            }
                            catch(Exception ex)
                            {
                                postErrors.Add($"Merge failed for group {groupNumber}: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"DXF export content failed: {failMsg}");
                    }
                }
            }
            finally
            {
                try { Directory.Delete(tempRoot, true); } catch { }
            }

            if (postErrors.Count > 0)
            {
                TaskDialog.Show("DXF Merge Errors", string.Join("\n", postErrors));
            }

            return success;
        }

        private string FindDxfForSheet(ViewSheet sheet, string folder)
        {
            var num = sheet.SheetNumber ?? "";
            // Typical Revit pattern: "Prefix-Sheet - <SheetNumber> - <SheetName>.dxf"
            foreach (var fp in Directory.EnumerateFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly))
            {
                var fn = Path.GetFileNameWithoutExtension(fp);
                if (fn.IndexOf($" - {num} - ", StringComparison.OrdinalIgnoreCase) >= 0)
                    return fp;
            }
            // Fallback
            foreach (var fp in Directory.EnumerateFiles(folder, "*.dxf", SearchOption.TopDirectoryOnly))
            {
                var fn = Path.GetFileNameWithoutExtension(fp);
                if (fn.IndexOf(num, StringComparison.OrdinalIgnoreCase) >= 0)
                    return fp;
            }
            return null;
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
