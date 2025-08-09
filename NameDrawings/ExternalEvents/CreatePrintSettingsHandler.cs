using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using EliteSheets.ExternalEvents;
using EliteSheets.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

public class CreatePrintSettingHandler : IExternalEventHandler
{
    public Document Doc { get; set; }
    public ViewSheet SheetToAnalyze { get; set; }  // The sheet used to define paper size
    public string PrintSettingName { get; set; } = "EliteSheetsZoom100";

    public ExternalEvent ExportEvent { get; set; } // Optional: used to trigger export after setup
    public ExportSheetsHandler ExportHandler { get; set; }
    public Action OnCompleted { get; set; }

    public void Execute(UIApplication app)
    {
        if (Doc == null || SheetToAnalyze == null)
        {
            OnCompleted?.Invoke();
            return;
        }

        try
        {
            using (Transaction tx = new Transaction(Doc, "Apply InSession Print Settings"))
            {
                tx.Start();

                var printManager = Doc.PrintManager;
                printManager.SelectNewPrintDriver("RK.Print");
                printManager.PrintRange = PrintRange.Select;
                printManager.PrintToFile = true;

                // Always use InSession setup
                var setup = printManager.PrintSetup;
                setup.CurrentPrintSetting = setup.InSession;

                var parameters = setup.CurrentPrintSetting.PrintParameters;
                ConfigurePrintParameters(parameters);

                TryAssignPaperSize(printManager, parameters, SheetToAnalyze);

                // Apply without saving
                printManager.Apply();

                tx.Commit();
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Print Error", $"Failed to apply print settings:\n{ex.Message}");
        }

        OnCompleted?.Invoke();
    }


    // Applies standard parameters to the print setup.
    private void ConfigurePrintParameters(PrintParameters parameters)
    {
        parameters.ZoomType = ZoomType.Zoom;
        parameters.Zoom = 100;
        parameters.HideCropBoundaries = true;
        parameters.HideReforWorkPlanes = true;
        parameters.HideScopeBoxes = true;
    }

    // Attempts to match and assign a paper size based on the given sheet's dimensions.
    private void TryAssignPaperSize(PrintManager printManager, PrintParameters parameters, ViewSheet sheet)
    {
        string isoLabel = PaperSizeHelper.GetPaperSizeLabel(sheet);

        if (string.IsNullOrEmpty(isoLabel))
            return;

        // Match strictly against standard ISO paper names
        PaperSize bestMatch = printManager.PaperSizes
            .Cast<PaperSize>()
            .FirstOrDefault(p =>
                string.Equals(p.Name.Trim(), isoLabel, StringComparison.OrdinalIgnoreCase) ||     // Exact match
                p.Name.Trim().StartsWith($"{isoLabel} ", StringComparison.OrdinalIgnoreCase) ||    // e.g. "A4 210 x 297 mm"
                p.Name.Trim().Equals($"{isoLabel}.Transverse", StringComparison.OrdinalIgnoreCase)); // Optional fallback

        if (bestMatch != null)
        {
            parameters.PaperSize = bestMatch;
            parameters.PaperPlacement = PaperPlacementType.Center;
        }
    }


    public string GetName() => "Create Print Setup with Zoom 100 and Paper Size";
}
