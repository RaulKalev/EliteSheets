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

            CleanupStaleTempFolders();

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
            var postErrors = new List<string>();

            foreach (var sheet in singles)
            {
                try
                {
                    if (dwgExporter.ExportSheet(sheet, out string failMsg))
                    {
                        success = true;
                    }
                    else if (!string.IsNullOrEmpty(failMsg))
                    {
                        postErrors.Add($"Sheet {sheet.SheetNumber}: {failMsg}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DWG export failed for {sheet.Name}: {ex.Message}");
                    postErrors.Add($"Sheet {sheet.SheetNumber}: {ex.Message}");
                }
            }

            if (postErrors.Count > 0)
            {
                TaskDialog.Show("DWG Export Errors", string.Join("\n", postErrors));
            }

            return success;
        }

        // --- PDF Logic ---

        private bool ExportAllPdfSheets(SheetGroupingService.PartitionResult partition)
        {
            bool success = false;
            var pdfExporter = new PdfExportService(Doc);
            var postErrors = new List<string>();

            // Singles
            foreach (var sheet in partition.Singles)
            {
                try
                {
                    if (pdfExporter.ExportSheetAsPdf(sheet, ExportPath))
                    {
                        success = true;
                    }
                    else
                    {
                        string target = Path.Combine(ExportPath, sheet.SheetNumber + ".pdf");
                        postErrors.Add($"Sheet {sheet.SheetNumber}: PDF export failed.{LockHint(target)}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"PDF export failed for {sheet.Name}: {ex.Message}");
                    postErrors.Add($"Sheet {sheet.SheetNumber}: {ex.Message}");
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
                    {
                        success = true;
                    }
                    else
                    {
                        string target = Path.Combine(ExportPath, outputName + ".pdf");
                        postErrors.Add($"Combined PDF failed for group {groupNumber}.{LockHint(target)}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Combined PDF export failed for group '{groupNumber}': {ex.Message}");
                    postErrors.Add($"Combined PDF failed for group {groupNumber}: {ex.Message}");
                }
            }

            if (postErrors.Count > 0)
            {
                TaskDialog.Show("PDF Export Errors", string.Join("\n", postErrors));
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

            // Temp folder lives in %TEMP%, not in the export folder: export folders are
            // often cloud-synced (Dropbox), and sync/antivirus locks on freshly written
            // temp files broke the in-place rewrite and left junk folders behind.
            string tempRoot = Path.Combine(GetScratchRoot(), "_tmp_dxf_merge_" + Guid.NewGuid().ToString("N"));
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
                            try { RetryOnSharingViolation(() => promoter.PromotePaperToModel(f)); }
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

                            if (IsFileLocked(outPath))
                            {
                                postErrors.Add($"Merge failed for group {groupNumber}: '{combinedName}.dxf' is open in another program. Close it and export again.");
                                continue;
                            }

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
                TryDeleteDirectory(tempRoot);
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

        // --- Temp file hygiene ---

        private static string GetScratchRoot() => Path.Combine(Path.GetTempPath(), "EliteSheets");

        /// <summary>
        /// Removes merge workspaces left behind by failed or crashed runs — both in the
        /// scratch root and in the export folder (where older versions created them).
        /// Only folders untouched for over an hour are removed, so an export running
        /// concurrently on another machine syncing into the same folder is never disturbed.
        /// </summary>
        private void CleanupStaleTempFolders()
        {
            SweepStaleTempFolders(GetScratchRoot());
            SweepStaleTempFolders(ExportPath);
        }

        private static void SweepStaleTempFolders(string root)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) return;

                var cutoff = DateTime.UtcNow.AddHours(-1);
                foreach (var dir in Directory.EnumerateDirectories(root, "_tmp_dxf_merge_*", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        if (Directory.GetLastWriteTimeUtc(dir) <= cutoff)
                            TryDeleteDirectory(dir);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Stale temp folder sweep failed in '{root}'.", ex);
            }
        }

        private static void TryDeleteDirectory(string path)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (!Directory.Exists(path)) return;

                    foreach (var f in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
                    {
                        try { File.SetAttributes(f, FileAttributes.Normal); } catch { }
                    }

                    Directory.Delete(path, true);
                    return;
                }
                catch (Exception ex)
                {
                    if (attempt == 2)
                    {
                        Logger.Log($"Could not delete temp folder '{path}'. It will be swept on the next export.", ex);
                        return;
                    }
                    System.Threading.Thread.Sleep(250);
                }
            }
        }

        /// <summary>
        /// Retries an IO action a few times — cloud-sync clients and antivirus scanners
        /// take short-lived locks on freshly written files.
        /// </summary>
        private static void RetryOnSharingViolation(Action action)
        {
            for (int attempt = 0; ; attempt++)
            {
                try { action(); return; }
                catch (IOException) when (attempt < 2)
                {
                    System.Threading.Thread.Sleep(200);
                }
            }
        }

        private static bool IsFileLocked(string path)
        {
            if (!File.Exists(path)) return false;
            try
            {
                using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                return false;
            }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return true; }
        }

        private static string LockHint(string targetPath)
        {
            return IsFileLocked(targetPath)
                ? $" The file '{Path.GetFileName(targetPath)}' is open in another program — close it and export again."
                : string.Empty;
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
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ExportPath,
                        UseShellExecute = true,
                        Verb = "open"
                    });
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
