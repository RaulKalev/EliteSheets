using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Newtonsoft.Json;
using EliteSheets.Commands;
using EliteSheets.Services;

namespace EliteSheets
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private RibbonPanel ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            // Define the custom tab name
            string tabName = "RK Tools";

            // Try to create the custom tab (avoid exception if it already exists)
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // Tab already exists; continue without throwing an error
            }

            // Create Ribbon Panel on the custom tab
            ribbonPanel = application.CreateOrSelectPanel(tabName, "Tools");

            // Ribbon text follows the saved language (Estonian until one is chosen)
            Loc.Load();

            // Create PushButton with embedded resource
            ribbonPanel.CreatePushButton<MainCommand>()
                .SetLargeImage("Assets/EliteSheets.tiff")
                .SetText(Loc.Get("RibbonText"))
                .SetToolTip(Loc.Get("RibbonTooltip"))
                .SetContextualHelp("https://raulkalev.github.io/rktools/");

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Trigger the update check
            ribbonPanel?.Remove();
            return Result.Succeeded;
        }

    }
}
