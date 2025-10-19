using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using EliteSheets.ExternalEvents;
using EliteSheets.Helpers;
using EliteSheets.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using WinForms = System.Windows.Forms;
using WpfComboBox = System.Windows.Controls.ComboBox;

namespace EliteSheets
{
    public partial class MainWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\EliteSheets\config.json";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd); // (kept for future use)

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); // (kept for future use)

        private const int SW_RESTORE = 9;

        #endregion

        #region Revit state / UI state

        private UIDocument _uiDoc;
        private Document _doc;
        private View _currentView;

        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isInitialized;

        public ObservableCollection<string> ViewTypes { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> ViewTemplates { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();

        #endregion

        #region External Events & Handlers

        private ExternalEvent _exportEvent;
        private ExportSheetsHandler _exportHandler;

        private ExternalEvent _createPrintSettingEvent;
        private CreatePrintSettingHandler _createPrintSettingHandler;

        private ExternalEvent _deletePrintSettingEvent;      // reserved for future use
        private ExternalEvent _EliteSheetsEvent;             // reserved for future use
        private ExternalEvent _generateEvent;                // reserved for future use

        #endregion

        #region Ctor / Init

        public MainWindow(UIDocument uiDoc, Document doc, View currentView)
        {
            InitializeComponent();

            _uiDoc = uiDoc;
            _doc = doc;
            _currentView = currentView;

            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // Window infrastructure
            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;

            // Window-level mouse hooks for resizing
            MouseMove += Window_MouseMove;
            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            // Theme + DataContext
            LoadThemeState();
            LoadTheme();
            DataContext = this;

            // Data
            _isInitialized = true;
            LoadExportPathForCurrentProject();
            LoadSheets();
            LoadDwgExportSetups();

            // External events
            _exportHandler = new ExportSheetsHandler();
            _exportEvent = ExternalEvent.Create(_exportHandler);

            _createPrintSettingHandler = new CreatePrintSettingHandler { Doc = _doc };
            _createPrintSettingEvent = ExternalEvent.Create(_createPrintSettingHandler);
        }

        #endregion

        #region Theme

        private void LoadTheme()
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            var themeUri = _isDarkMode
                ? $"pack://application:,,,/{assemblyName};component/UI/Themes/DarkTheme.xaml"
                : $"pack://application:,,,/{assemblyName};component/UI/Themes/LightTheme.xaml";

            try
            {
                var resourceDict = new ResourceDictionary { Source = new Uri(themeUri, UriKind.Absolute) };
                Resources.MergedDictionaries.Clear();
                Resources.MergedDictionaries.Add(resourceDict);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load theme: {ex.Message}\nTheme URI: {themeUri}", "Theme Load Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = ThemeToggleButton.IsChecked == true;
            LoadTheme();

            // Ensure the icon reflects the current state
            var icon = ThemeToggleButton?.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                       as MaterialDesignThemes.Wpf.PackIcon;
            if (icon != null)
            {
                icon.Kind = _isDarkMode
                    ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                    : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
            }
        }

        private void LoadThemeState()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (config != null &&
                        config.TryGetValue("IsDarkMode", out var isDarkModeObj) &&
                        isDarkModeObj is bool isDark)
                    {
                        _isDarkMode = isDark;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load theme state: {ex.Message}", "Load Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Reflect the state in the toggle (if template is ready)
            ThemeToggleButton.IsChecked = _isDarkMode;
            var icon = ThemeToggleButton?.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                       as MaterialDesignThemes.Wpf.PackIcon;
            if (icon != null)
            {
                icon.Kind = _isDarkMode
                    ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                    : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
            }
        }

        private void SaveThemeState()
        {
            try
            {
                var config = new Dictionary<string, object>();

                if (File.Exists(ConfigFilePath))
                {
                    var existingJson = File.ReadAllText(ConfigFilePath);
                    config = JsonConvert.DeserializeObject<Dictionary<string, object>>(existingJson)
                             ?? new Dictionary<string, object>();
                }

                config["IsDarkMode"] = _isDarkMode;

                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFilePath));
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save theme state: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Config: Export path per project

        private void SaveExportPathForCurrentProject(string exportPath)
        {
            try
            {
                var projectName = Path.GetFileName(_doc.PathName);
                if (string.IsNullOrWhiteSpace(projectName)) return;

                var config = new Dictionary<string, object>();

                if (File.Exists(ConfigFilePath))
                {
                    var existingJson = File.ReadAllText(ConfigFilePath);
                    config = JsonConvert.DeserializeObject<Dictionary<string, object>>(existingJson)
                             ?? new Dictionary<string, object>();
                }

                Dictionary<string, string> exportPaths;
                if (config.TryGetValue("ExportPaths", out object rawPaths) &&
                    rawPaths is Newtonsoft.Json.Linq.JObject jObj)
                {
                    exportPaths = jObj.ToObject<Dictionary<string, string>>();
                }
                else
                {
                    exportPaths = new Dictionary<string, string>();
                }

                exportPaths[projectName] = exportPath;
                config["ExportPaths"] = exportPaths;

                if (!config.ContainsKey("IsDarkMode"))
                    config["IsDarkMode"] = _isDarkMode;

                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Save Error", $"Failed to save export path:\n{ex.Message}");
            }
        }

        private void LoadExportPathForCurrentProject()
        {
            try
            {
                if (!File.Exists(ConfigFilePath)) return;

                var json = File.ReadAllText(ConfigFilePath);
                var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (config == null || !config.ContainsKey("ExportPaths")) return;

                if (config["ExportPaths"] is Newtonsoft.Json.Linq.JObject jObj)
                {
                    var exportPaths = jObj.ToObject<Dictionary<string, string>>();
                    var projectName = Path.GetFileName(_doc.PathName);

                    if (!string.IsNullOrEmpty(projectName) &&
                        exportPaths.TryGetValue(projectName, out var savedPath))
                    {
                        ExportPathTextBox.Text = savedPath;
                    }
                }
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Load Error", $"Failed to load export path:\n{ex.Message}");
            }
        }

        #endregion

        #region Data loading

        private void LoadSheets()
        {
            Sheets.Clear();

            var collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>();

            foreach (var sheet in collector)
            {
                // first placed view name on sheet (optional info)
                var viewName = "";
                var viewports = new FilteredElementCollector(_doc, sheet.Id)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>();

                var firstViewport = viewports.FirstOrDefault();
                if (firstViewport != null)
                {
                    var view = _doc.GetElement(firstViewport.ViewId) as View;
                    if (view != null) viewName = view.Name;
                }

                var item = new SheetItem
                {
                    Name = sheet.Name,
                    Number = sheet.SheetNumber,
                    ViewName = viewName,
                    Id = sheet.Id,
                    IsChecked = false
                };

                // size label (existing)
                var outline = sheet.Outline;
                if (outline != null)
                {
                    var width = UnitUtils.ConvertFromInternalUnits(outline.Max.U - outline.Min.U, UnitTypeId.Millimeters);
                    var height = UnitUtils.ConvertFromInternalUnits(outline.Max.V - outline.Min.V, UnitTypeId.Millimeters);
                    item.SheetSize = PaperSizeHelper.GetPaperSizeLabel(width, height);
                }
                else
                {
                    item.SheetSize = "Unknown";
                }

                // NEW: latest version / revision label
                string versionText = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION)?.AsString();
                if (string.IsNullOrWhiteSpace(versionText))
                {
                    // Fallback: take the highest SequenceNumber from Revisions placed on this sheet
                    var revIds = sheet.GetAllRevisionIds();
                    if (revIds != null && revIds.Count > 0)
                    {
                        var revisions = revIds
                            .Select(id => _doc.GetElement(id) as Revision)
                            .Where(r => r != null);

                        var latest = revisions
                            .OrderByDescending(r => r.SequenceNumber)
                            .FirstOrDefault();

                        versionText = latest?.RevisionNumber
                                      ?? latest?.SequenceNumber.ToString();
                    }
                }
                item.Version = string.IsNullOrWhiteSpace(versionText) ? "-" : versionText;

                Sheets.Add(item);

            }
        }

        private void LoadDwgExportSetups()
        {
            try
            {
                var setups = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .OrderBy(s => s.Name)
                    .ToList();

                DwgExportComboBox.ItemsSource = setups;
                if (setups.Any())
                    DwgExportComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Error", $"Failed to load DWG export setups: {ex.Message}");
            }
        }

        #endregion

        #region Button / UI handlers
        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // If you later add a "ViewTypeComboBox", you can still read it unambiguously:
                // var viewType = (FindName("ViewTypeComboBox") as WpfComboBox)?.SelectedItem as string;

                LoadSheets();               // refresh the grid’s backing collection
                SheetsDataGrid?.Items.Refresh();

                // Optional: also refresh DWG setups so the UI stays in sync
                LoadDwgExportSetups();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Reload Error", $"Failed to reload data: {ex.Message}");
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dlg = new WinForms.FolderBrowserDialog())
                {
                    dlg.Description = "Select export folder";
                    dlg.ShowNewFolderButton = true;

                    var initial = ExportPathTextBox.Text;
                    dlg.SelectedPath = (!string.IsNullOrWhiteSpace(initial) && Directory.Exists(initial))
                        ? initial
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    if (dlg.ShowDialog() == WinForms.DialogResult.OK)
                    {
                        ExportPathTextBox.Text = dlg.SelectedPath;
                        SaveExportPathForCurrentProject(dlg.SelectedPath);
                    }
                }
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Browse Error", ex.Message);
            }
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // 1) validate selection + export options
            var selected = Sheets.Where(s => s.IsChecked).ToList();
            if (!selected.Any())
            {
                RevitTaskDialog.Show("Error", "No sheets selected.");
                return;
            }

            var exportPath = ExportPathTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(exportPath) || !Directory.Exists(exportPath))
            {
                RevitTaskDialog.Show("Error", "Please select a valid export folder.");
                return;
            }

            var exportDwg = (LogicalTreeHelper.FindLogicalNode(this, "DwgExportCheckbox") as CheckBox)?.IsChecked == true;
            var exportPdf = (LogicalTreeHelper.FindLogicalNode(this, "PdfExportCheckbox") as CheckBox)?.IsChecked == true;

            if (!exportDwg && !exportPdf)
            {
                RevitTaskDialog.Show("Info", "Neither DWG nor PDF export is selected.");
                return;
            }

            var exportSetup = DwgExportComboBox.SelectedItem as ExportDWGSettings;
            if (exportDwg && exportSetup == null)
            {
                RevitTaskDialog.Show("Error", "Please select a DWG export setup.");
                return;
            }

            // 2) validate filenames
            var forbidden = Path.GetInvalidFileNameChars();
            var invalidSheets = selected
                .Where(s => s.Number.IndexOfAny(forbidden) >= 0)
                .Select(s => new
                {
                    Sheet = s,
                    Invalid = new string(s.Number.Where(c => forbidden.Contains(c)).Distinct().ToArray())
                })
                .ToList();

            if (invalidSheets.Any())
            {
                var message = "Järgnevatel lehtedel on mittesobivad märgid nende lehenumbris:\n\n" +
                              string.Join("\n", invalidSheets.Select(i => $"• \"{i.Sheet.Number}\" → {i.Invalid}")) +
                              "\n\nWindows ei luba järgmisi märke failinimedes:\n" +
                              string.Join(" ", forbidden.Select(c => $"'{c}'")) +
                              "\n\nKas soovid eksportida ülejäänud lehed?";

                var result = RevitTaskDialog.Show("Mittesobivad joonise numbrid", message,
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    TaskDialogResult.No);

                if (result == TaskDialogResult.No) return;

                selected = selected.Except(invalidSheets.Select(i => i.Sheet)).ToList();
                if (!selected.Any())
                {
                    RevitTaskDialog.Show("No Valid Sheets", "All selected sheets have invalid characters. Nothing to export.");
                    return;
                }
            }

            // 3) resolve ViewSheet elements
            var sheetElements = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(vs => selected.Any(s => s.Id == vs.Id))
                .ToList();

            if (!sheetElements.Any())
            {
                RevitTaskDialog.Show("Error", "Could not resolve selected sheets in the document.");
                return;
            }

            // 4) raise export event
            _exportHandler.UiDoc = _uiDoc;
            _exportHandler.Doc = _doc;
            _exportHandler.SheetsToExport = sheetElements;
            _exportHandler.ExportPath = exportPath;
            _exportHandler.ExportSetupName = exportSetup?.Name;
            _exportHandler.ExportPdf = exportPdf;
            _exportHandler.ExportDwg = exportDwg;

            _exportEvent.Raise();
        }

        private void Checkbox_Click(object sender, RoutedEventArgs e)
        {
            // Multi-select toggle propagation
            if (SheetsDataGrid.SelectedItems.Count <= 1) return;

            var checkBox = sender as CheckBox;
            var clickedItem = checkBox?.DataContext as SheetItem;
            if (clickedItem == null) return;

            var newState = checkBox.IsChecked == true;
            foreach (var selected in SheetsDataGrid.SelectedItems)
            {
                var item = selected as SheetItem;
                if (item != null && !ReferenceEquals(item, clickedItem))
                    item.IsChecked = newState;
            }

            SheetsDataGrid.Items.Refresh();
        }

        private void CheckAllBox_Click(object sender, RoutedEventArgs e)
        {
            var headerCheckbox = sender as CheckBox;
            if (headerCheckbox == null) return;

            var newState = headerCheckbox.IsChecked == true;
            foreach (var item in Sheets) item.IsChecked = newState;

            SheetsDataGrid.Items.Refresh();
        }

        private void SheetsDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Ignore if multi-select modifiers are used
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl) ||
                Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                return;

            var dataGrid = sender as DataGrid;
            if (dataGrid == null) return;

            var hit = VisualTreeHelper.HitTest(dataGrid, e.GetPosition(dataGrid));
            var current = hit?.VisualHit;

            // If click is on a row, header or scrollbar, don't clear selection
            while (current != null)
            {
                if (current is DataGridRow || current is ScrollBar || current is DataGridColumnHeader)
                    return;
                current = VisualTreeHelper.GetParent(current);
            }

            // Otherwise, clear selection (delayed so checkbox clicks still toggle)
            dataGrid.Dispatcher.BeginInvoke(new Action(() => dataGrid.UnselectAll()),
                DispatcherPriority.Input);
        }

        #endregion

        #region Window chrome / resize handlers

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
        private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

        private void LeftEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void RightEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void BottomEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNS;
        private void Edge_MouseLeave(object sender, MouseEventArgs e) => Cursor = Cursors.Arrow;
        private void BottomLeftCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNESW;
        private void BottomRightCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNWSE;

        private void Window_MouseMove(object sender, MouseEventArgs e) => _windowResizer.ResizeWindow(e);
        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) => _windowResizer.StopResizing();
        private void LeftEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Left);
        private void RightEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Right);
        private void BottomEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Bottom);
        private void BottomLeftCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomLeft);
        private void BottomRightCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomRight);

        #endregion

        #region Cleanup / disposal

        private static void DisposeExternalEvent(ref ExternalEvent ev)
        {
            if (ev != null)
            {
                try { ev.Dispose(); } catch { /* ignore on shutdown */ }
                ev = null;
            }
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            try
            {
                Closed -= MainWindow_Closed;

                SaveThemeState();

                if (SheetsDataGrid != null) SheetsDataGrid.ItemsSource = null;

                if (ViewTypes != null) ViewTypes.Clear();
                if (ViewTemplates != null) ViewTemplates.Clear();
                if (Sheets != null) Sheets.Clear();

                DisposeExternalEvent(ref _exportEvent);
                DisposeExternalEvent(ref _createPrintSettingEvent);
                DisposeExternalEvent(ref _deletePrintSettingEvent);
                DisposeExternalEvent(ref _EliteSheetsEvent);
                DisposeExternalEvent(ref _generateEvent);

                _exportHandler = null;
                _createPrintSettingHandler = null;

                _uiDoc = null;
                _doc = null;
                _currentView = null;

                var disposableResizer = _windowResizer as IDisposable;
                if (disposableResizer != null) { try { disposableResizer.Dispose(); } catch { } }
            }
            catch
            {
                // swallow – app is closing
            }
        }

        #endregion

        #region Misc

        private void TitleBar_Loaded(object sender, RoutedEventArgs e)
        {
            // keep hook if you add logic later
        }

        #endregion
    }
}
