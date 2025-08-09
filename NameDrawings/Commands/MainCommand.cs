using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace EliteSheets.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class MainCommand : IExternalCommand
    {
        private static MainWindow _window; // single instance per Revit process

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // If window already exists, just surface it
                if (_window != null && _window.IsLoaded)
                {
                    var hwnd = new WindowInteropHelper(_window).Handle;
                    if (_window.WindowState == WindowState.Minimized)
                        ShowWindow(hwnd, SW_RESTORE);

                    // Nudge to front
                    _window.Activate();
                    _window.Focus();
                    SetForegroundWindow(hwnd);
                    return Result.Succeeded;
                }

                // Create new instance
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                View currentView = doc.ActiveView;

                _window = new MainWindow(uiDoc, doc, currentView);
                _window.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                var owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new WindowInteropHelper(_window) { Owner = owner };

                // When closed, release the static reference
                _window.Closed += (s, e) => { _window = null; };

                _window.Show(); // modeless
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
