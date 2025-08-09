using Autodesk.Revit.DB;
using EliteSheets.Helpers;
using System;
using System.Collections.Generic;
using System.IO;

namespace EliteSheets.Exports
{
    public class PdfExportService
    {
        private readonly Document _doc;

        public PdfExportService(Document doc) => _doc = doc;

        public bool ExportSheetAsPdf(ViewSheet sheet, string exportFolder)
        {
            if (sheet == null) return false;

            Directory.CreateDirectory(exportFolder);

            // Build a ViewId list with just this sheet
            var toExport = new List<ElementId> { sheet.Id };

            // Configure PDF options
            using (var opts = new PDFExportOptions())
            {
                // one PDF per sheet
                opts.Combine = false;

                // 100% zoom, centered, hide clutter (same intent as before)
                opts.ZoomType = ZoomType.Zoom;
                opts.ZoomPercentage = 100;
                opts.PaperPlacement = PaperPlacementType.Center;
                opts.HideCropBoundaries = true;
                opts.HideReferencePlane = true;
                opts.HideScopeBoxes = true;
                opts.StopOnError = false;

                // Optional quality (vector where possible)
                opts.AlwaysUseRaster = false;

                // Paper match: not strictly required (Revit uses the sheet),
                // but we keep your helper in case you extend sizing later.
                var isoLabel = PaperSizeHelper.GetPaperSizeLabel(sheet);
                // (No direct paper enum mapping needed; leave PaperFormat as default)

                // Name each output file by SHEET_NUMBER
                // Build a naming rule: a single “Sheet Number” token, no prefix/suffix.
                var rule = new List<TableCellCombinedParameterData>();
                var part = TableCellCombinedParameterData.Create();
                part.ParamId = new ElementId(BuiltInParameter.SHEET_NUMBER);
                part.CategoryId = new ElementId(BuiltInCategory.OST_Sheets);
                part.Separator = "";
                part.Prefix = "";
                part.Suffix = "";
                rule.Add(part);
                opts.SetNamingRule(rule);

                // Export to the chosen folder
                // Revit will create "<SheetNumber>.pdf" directly in exportFolder.
                bool ok = _doc.Export(exportFolder, toExport, opts);
                return ok;
            }
        }
        public bool ExportCombinedPdf(List<ViewSheet> sheets, string exportFolder, string outputNameWithoutExt)
        {
            if (sheets == null || sheets.Count == 0) return false;

            Directory.CreateDirectory(exportFolder);

            // Build a safe base filename and target path
            string safeBase = Sanitize(outputNameWithoutExt);
            if (string.IsNullOrWhiteSpace(safeBase)) safeBase = "Combined";
            string target = Path.Combine(exportFolder, safeBase + ".pdf");

            // Collect sheet ids
            var toExport = new List<ElementId>();
            foreach (var s in sheets) toExport.Add(s.Id);

            // Configure PDF options for a single combined file
            using (var opts = new PDFExportOptions())
            {
                opts.Combine = true;            // one PDF for all sheets
                opts.FileName = safeBase;       // Revit writes "<safeBase>.pdf" into exportFolder

                // Keep your existing visual prefs (most apply to combined too)
                opts.ZoomType = ZoomType.Zoom;
                opts.ZoomPercentage = 100;
                opts.PaperPlacement = PaperPlacementType.Center;
                opts.HideCropBoundaries = true;
                opts.HideReferencePlane = true;
                opts.HideScopeBoxes = true;
                opts.StopOnError = false;
                opts.AlwaysUseRaster = false;

                // No naming rule when Combine=true (FileName is used instead)
                bool ok = _doc.Export(exportFolder, toExport, opts);
                // Confirm the expected file exists
                return ok && File.Exists(target);
            }
        }

        // helper
        private static string Sanitize(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Output";
            var invalid = Path.GetInvalidFileNameChars();
            var arr = name.ToCharArray();
            for (int i = 0; i < arr.Length; i++)
                if (Array.IndexOf(invalid, arr[i]) >= 0) arr[i] = '_';
            return new string(arr);
        }

    }
}
