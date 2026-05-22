using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace EliteSheets.Exports
{
    public class DwgExportService
    {
        private readonly Document _doc;
        private readonly DWGExportOptions _options;
        private readonly string _exportFolder;

        public DwgExportService(Document doc, DWGExportOptions options, string exportFolder)
        {
            _doc = doc;
            _options = options;
            _exportFolder = exportFolder;
        }

        public bool ExportSheet(ViewSheet sheet, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                ICollection<ElementId> sheetIds = new List<ElementId> { sheet.Id };
                string subfolderName = sheet.SheetNumber;

                Directory.CreateDirectory(_exportFolder);

                return _doc.Export(_exportFolder, subfolderName, sheetIds, _options);
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Debug.WriteLine($"DWG export failed for {sheet.SheetNumber}: {ex}");
                return false;
            }
        }
    }
}
