using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace EliteSheets.Services
{
    /// <summary>
    /// Service for exporting views/sheets to DXF.
    /// Usage: pass a list of View/Sheet ElementIds and an output folder.
    /// </summary>
    public class DxfExportService
    {
        /// <summary>
        /// Exports the given views/sheets to DXF.
        /// Revit creates one DXF per view/sheet using the baseFileName as prefix.
        /// </summary>
        /// <param name="doc">Active Revit document.</param>
        /// <param name="viewOrSheetIds">Views and/or sheets to export.</param>
        /// <param name="outputFolder">Destination folder (created if missing).</param>
        /// <param name="baseFileName">Base file name (prefix) used by Revit.</param>
        /// <param name="predefinedSetupName">Optional: name of a DWG/DXF export setup to reuse.</param>
        /// <param name="openFolderWhenDone">Open output folder on success.</param>
        /// <param name="failureMessage">Filled on failure.</param>
        /// <returns>true on success, false otherwise.</returns>
        public bool Export(
            Document doc,
            IList<ElementId> viewOrSheetIds,
            string outputFolder,
            string baseFileName,
            string predefinedSetupName,
            bool openFolderWhenDone,
            out string failureMessage)
        {
            failureMessage = string.Empty;

            try
            {
                if (doc == null)
                {
                    failureMessage = "No active document.";
                    return false;
                }
                if (viewOrSheetIds == null || viewOrSheetIds.Count == 0)
                {
                    failureMessage = "No views or sheets selected.";
                    return false;
                }
                if (string.IsNullOrWhiteSpace(outputFolder))
                {
                    failureMessage = "Output folder is empty.";
                    return false;
                }

                Directory.CreateDirectory(outputFolder);

                var options = BuildOptions(doc, predefinedSetupName);

                // One DXF per view/sheet. Names are based on baseFileName + view/sheet info.
                bool ok = doc.Export(outputFolder, baseFileName ?? "DXF_Export", viewOrSheetIds, options);

                if (!ok)
                {
                    failureMessage = "Revit reported that the DXF export failed. Verify the export setup and selected views/sheets.";
                    return false;
                }

                if (openFolderWhenDone && Directory.Exists(outputFolder))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = outputFolder,
                            UseShellExecute = true
                        });
                    }
                    catch
                    {
                        // Non-fatal if we can't open the folder
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                failureMessage = ex.Message;
                return false;
            }
        }

        private static DXFExportOptions BuildOptions(Document doc, string setupName)
        {
            DXFExportOptions options = null;

            // Try to use a predefined DWG/DXF export setup from the model, if provided
            if (!string.IsNullOrWhiteSpace(setupName))
            {
                try
                {
                    options = DXFExportOptions.GetPredefinedOptions(doc, setupName);
                }
                catch
                {
                    // Ignore if setup not found or not applicable; fall back to defaults
                    options = null;
                }
            }

            if (options == null)
                options = new DXFExportOptions();

            // Optional: adjust defaults here if you want
            // options.SharedCoords = true;
            // options.TargetUnit = ExportUnit.Millimeter;
            // options.FileVersion = ACADVersion.R2018; // if available in your Revit version

            return options;
        }
    }
}
