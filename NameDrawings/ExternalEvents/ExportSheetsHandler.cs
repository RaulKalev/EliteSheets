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
        /// <summary>Optional PDF destination; falls back to <see cref="ExportPath"/> when empty.</summary>
        public string PdfExportPath { get; set; }
        /// <summary>Optional DWG/DXF destination; falls back to <see cref="ExportPath"/> when empty.</summary>
        public string DwgExportPath { get; set; }
        public string ExportSetupName { get; set; }
        public bool ExportPdf { get; set; } = true;
        public bool ExportDwg { get; set; } = true;
        public bool ExportDxf { get; set; } = false;
        public string TemplateDxfPath { get; set; }
        /// <summary>When true, merged group DXFs are converted to DWG with ODA File Converter.</summary>
        public bool ConvertMergedToDwg { get; set; }
        /// <summary>Plugin window that result dialogs are shown on top of.</summary>
        public System.Windows.Window OwnerWindow { get; set; }
        private System.Windows.Window ActiveOwner => OwnerWindow != null && OwnerWindow.IsLoaded ? OwnerWindow : null;

        private readonly SheetGroupingService _groupingService = new SheetGroupingService();

        private string PdfFolder => string.IsNullOrWhiteSpace(PdfExportPath) ? ExportPath : PdfExportPath;
        private string CadFolder => string.IsNullOrWhiteSpace(DwgExportPath) ? ExportPath : DwgExportPath;

        public void Execute(UIApplication app)
        {
            if (Doc == null || UiDoc == null || SheetsToExport == null)
                return;
            if (ExportPdf && string.IsNullOrWhiteSpace(PdfFolder))
                return;
            if ((ExportDwg || ExportDxf) && string.IsNullOrWhiteSpace(CadFolder))
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
            var dwgExporter = new DwgExportService(Doc, options, CadFolder);
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
                        postErrors.Add(Loc.Format("SheetError", sheet.SheetNumber, failMsg));
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DWG export failed for {sheet.Name}: {ex.Message}");
                    postErrors.Add(Loc.Format("SheetError", sheet.SheetNumber, ex.Message));
                }
            }

            if (postErrors.Count > 0)
            {
                ShowMessage(Loc.Get("DwgExportErrorsTitle"), string.Join("\n", postErrors), ThemedDialogKind.Error);
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
                    if (pdfExporter.ExportSheetAsPdf(sheet, PdfFolder))
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
                    if (pdfExporter.ExportCombinedPdf(orderedSheets, PdfFolder, outputName))
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

            string outFolder = CadFolder;
            Directory.CreateDirectory(outFolder);

            // Snapshot existing DXFs to identify new ones
            var pre = new HashSet<string>(
                Directory.EnumerateFiles(outFolder, "*.dxf", SearchOption.TopDirectoryOnly),
                StringComparer.OrdinalIgnoreCase);

            if (!dxfExporter.Export(Doc, ids, outFolder,"DXF_Sheets", ExportSetupName, false, out string failureMsg))
            {
                 Debug.WriteLine($"DXF export (singles) failed: {failureMsg}");
            }
            else
            {
                 success = true;
                 
                 // Identify newly created files
                 var newFiles = Directory.EnumerateFiles(outFolder, "*.dxf", SearchOption.TopDirectoryOnly)
                                         .Where(p => !pre.Contains(p))
                                         .ToList();
                 
                 if (newFiles.Count == 0)
                 {
                     var cutoff = DateTime.UtcNow.AddMinutes(-2);
                     newFiles = Directory.EnumerateFiles(outFolder, "*.dxf", SearchOption.TopDirectoryOnly)
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
                ShowMessage(Loc.Get("DxfMergeTitle"), Loc.Get("TemplateMissing"), ThemedDialogKind.Warning);
                return false;
            }

            // Temp folder
            string tempRoot = Path.Combine(CadFolder, "_tmp_dxf_merge_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);

            var dxfExporter = new EliteSheets.Services.DxfExportService();
            var promoter = new EliteSheets.Services.DxfPaperToModelPromoter();
            var merger = new EliteSheets.Services.DxfMergeService();

            // When converting, merged DXFs go to a temp folder and ODA writes the DWGs to CadFolder
            string odaPath = ConvertMergedToDwg ? OdaConverterService.FindConverter() : null;
            if (ConvertMergedToDwg && odaPath == null)
                postErrors.Add(Loc.Get("OdaMissingSavedDxf"));
            string mergedFolder = odaPath != null ? Path.Combine(tempRoot, "_merged") : CadFolder;
            if (odaPath != null) Directory.CreateDirectory(mergedFolder);

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
                                    postErrors.Add(Loc.Format("DxfNotFoundForMerge", s.SheetNumber));
                            }

                            if (sourcePaths.Count == 0) continue;

                            string combinedName = _groupingService.BuildCombinedFileName(orderedSheets.First().SheetNumber, groupNumber);
                            string outPath = Path.Combine(mergedFolder, combinedName + ".dxf");

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
                                postErrors.Add(Loc.Format("MergeFailed", groupNumber, ex.Message));
                            }
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"DXF export content failed: {failMsg}");
                    }
                }

                if (odaPath != null)
                    ConvertMergedDxfsToDwg(odaPath, mergedFolder, postErrors);
            }
            finally
            {
                try { Directory.Delete(tempRoot, true); } catch { }
            }

            if (postErrors.Count > 0)
            {
                ShowMessage(Loc.Get("DxfMergeErrorsTitle"), string.Join("\n", postErrors), ThemedDialogKind.Error);
            }

            return success;
        }

        /// <summary>
        /// Converts merged DXFs to DWG in CadFolder. Any file that fails to convert is copied as DXF instead.
        /// </summary>
        private void ConvertMergedDxfsToDwg(string odaPath, string mergedFolder, List<string> postErrors)
        {
            var dxfs = Directory.GetFiles(mergedFolder, "*.dxf", SearchOption.TopDirectoryOnly);
            if (dxfs.Length == 0) return;

            var converter = new OdaConverterService();
            if (!converter.ConvertFolder(odaPath, mergedFolder, CadFolder, out string odaError))
                postErrors.Add(Loc.Format("DwgConversionFailed", odaError));

            foreach (var dxf in dxfs)
            {
                string dwg = Path.Combine(CadFolder, Path.GetFileNameWithoutExtension(dxf) + ".dwg");
                if (File.Exists(dwg) && File.GetLastWriteTimeUtc(dwg) >= File.GetLastWriteTimeUtc(dxf))
                    continue;

                // Fall back to delivering the DXF so the merged drawing isn't lost
                try
                {
                    File.Copy(dxf, Path.Combine(CadFolder, Path.GetFileName(dxf)), true);
                    postErrors.Add(Loc.Format("NotConvertedSavedDxf", Path.GetFileName(dxf)));
                }
                catch (Exception ex)
                {
                    postErrors.Add($"{Path.GetFileName(dxf)}: {ex.Message}");
                }
            }
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
                bool openFolder = ActiveOwner != null
                    ? ThemedMessageDialog.AskYesNo(ActiveOwner, Loc.Get("ExportDoneTitle"),
                        Loc.Get("ExportDoneMessage"), ThemedDialogKind.Success)
                    : TaskDialog.Show(
                        Loc.Get("ExportDoneTitle"),
                        Loc.Get("ExportDoneMessage"),
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                        TaskDialogResult.No) == TaskDialogResult.Yes;

                if (openFolder)
                {
                    foreach (var folder in GetUsedFolders())
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = folder,
                            UseShellExecute = true,
                            Verb = "open"
                        });
                    }
                }
            }
            else
            {
                ShowMessage(Loc.Get("ExportFailedTitle"), Loc.Get("ExportFailed"), ThemedDialogKind.Error);
            }
        }

        /// <summary>
        /// Shows a themed dialog on top of the plugin window, or a Revit TaskDialog when no window is attached.
        /// </summary>
        private void ShowMessage(string title, string message, ThemedDialogKind kind)
        {
            if (ActiveOwner != null)
                ThemedMessageDialog.Show(ActiveOwner, title, message, kind);
            else
                TaskDialog.Show(title, message);
        }

        private IEnumerable<string> GetUsedFolders()
        {
            var folders = new List<string>();
            if (ExportDwg || ExportDxf) folders.Add(CadFolder);
            if (ExportPdf) folders.Add(PdfFolder);

            return folders
                .Where(f => !string.IsNullOrWhiteSpace(f) && Directory.Exists(f))
                .Select(f => Path.GetFullPath(f).TrimEnd(Path.DirectorySeparatorChar))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        public string GetName() => "Export Sheets Handler";
    }
}
