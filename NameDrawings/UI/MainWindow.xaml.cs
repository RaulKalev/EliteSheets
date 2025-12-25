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
using Microsoft.Win32; // OpenFileDialog
using System.ComponentModel;
using System.Windows.Data;
using System.Collections;
using System.Text.RegularExpressions;

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
        private ICollectionView _sheetsView;

        #endregion

        #region Revit state / UI state

        private UIDocument _uiDoc;
        private Document _doc;
        private View _currentView;

        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isInitialized;
        private string _templateDxfPath = string.Empty;

        // Background scanning state
        private UIApplication _uiapp;
        private Queue<SheetItem> _sheetsToScan = new Queue<SheetItem>();
        private const int SheetsPerIdleTick = 20;

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
            Loaded += (s, e) =>
            {
                // run after the window is visible and Revit is idle
                Dispatcher.BeginInvoke(new Action(EnsureTemplatePathConfigured),
                    System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            };
            LoadTheme();
            DataContext = this;

            // Data
            _isInitialized = true;
            LoadExportPathForCurrentProject();
            LoadSheets();

            _sheetsView = CollectionViewSource.GetDefaultView(Sheets);
            _sheetsView.Filter = SheetsFilter;

            // Default sort on launch: Sheet Number (natural)
            var listView = _sheetsView as ListCollectionView;
            if (listView != null)
            {
                // IMPORTANT: do not use SortDescriptions together with CustomSort
                listView.SortDescriptions.Clear();
                listView.CustomSort = new SheetNumberComparer();
            }
            else
            {
                // fallback: basic string sort if view is not a ListCollectionView (rare)
                _sheetsView.SortDescriptions.Clear();
                _sheetsView.SortDescriptions.Add(new SortDescription(nameof(SheetItem.Number), ListSortDirection.Ascending));
            }

            _sheetsView.Refresh();



            UpdateClearButtonState();

            LoadDwgExportSetups();


            // External events
            _exportHandler = new ExportSheetsHandler();
            _exportEvent = ExternalEvent.Create(_exportHandler);

            _createPrintSettingHandler = new CreatePrintSettingHandler { Doc = _doc };
            _createPrintSettingEvent = ExternalEvent.Create(_createPrintSettingHandler);

            // Subscribe to Idling for background paper size scanning
            _uiapp = new UIApplication(doc.Application);
            _uiapp.Idling += OnIdling_ScanPaperSizes;
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
            SaveThemeState(); // <-- add this

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

                    if (config != null)
                    {
                        if (config.TryGetValue("IsDarkMode", out var isDarkModeObj) && isDarkModeObj is bool isDark)
                            _isDarkMode = isDark;

                        if (config.TryGetValue("TemplateDxfPath", out var templateObj))
                            _templateDxfPath = templateObj?.ToString() ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load theme/config: {ex.Message}", "Load Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // reflect UI
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

                // preserve/export paths and other keys; just update these two
                config["IsDarkMode"] = _isDarkMode;
                if (!string.IsNullOrWhiteSpace(_templateDxfPath))
                    config["TemplateDxfPath"] = _templateDxfPath;

                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFilePath));
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void EnsureTemplatePathConfigured()
        {
            if (!IsLoaded)
            {
                Dispatcher.BeginInvoke(new Action(EnsureTemplatePathConfigured),
                    System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                return;
            }
            // already set and exists → nothing to do
            if (!string.IsNullOrWhiteSpace(_templateDxfPath) && File.Exists(_templateDxfPath))
                return;

            var user = Environment.UserName;
            string suggestedFolder = $@"C:\Users\{user}\EULE Dropbox\0_EULE  Team folder (kogu kollektiiv)\02_EULE REVIT TEMPLATE";


            MessageBox.Show(
                "Palun vali DXF template KilbiTemplate.dxf (tingimata .dxf). " +
                "Soovitatav asukoht avatakse dialoogis automaatselt." +
                "\nFail asub kaustas:  \\EULE Dropbox\\0_EULE  Team folder (kogu kollektiiv)\\02_EULE REVIT TEMPLATE ",
                "Vajalik seadistus", MessageBoxButton.OK, MessageBoxImage.Information);

            var dlg = new OpenFileDialog
            {
                Title = "Vali mall (DXF)",
                Filter = "DXF fail (*.dxf)|*.dxf",
                CheckFileExists = true,
                Multiselect = false,
                InitialDirectory = Directory.Exists(suggestedFolder)
                    ? suggestedFolder
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                FileName = "KilbiTemplate.dxf"
            };

            if (dlg.ShowDialog() == true)
            {
                // ...
            }
            _templateDxfPath = dlg.FileName;
            SaveThemeState();
        }
        private sealed class SheetNumberComparer : IComparer
        {
            public int Compare(object x, object y)
            {
                var a = x as SheetItem;
                var b = y as SheetItem;

                var an = a?.Number ?? string.Empty;
                var bn = b?.Number ?? string.Empty;

                return CompareNatural(an, bn);
            }

            private static int CompareNatural(string a, string b)
            {
                if (ReferenceEquals(a, b)) return 0;
                if (a == null) return -1;
                if (b == null) return 1;

                var ax = Tokenize(a);
                var bx = Tokenize(b);

                int n = Math.Min(ax.Count, bx.Count);
                for (int i = 0; i < n; i++)
                {
                    var ta = ax[i];
                    var tb = bx[i];

                    bool na = int.TryParse(ta, out int ia);
                    bool nb = int.TryParse(tb, out int ib);

                    int cmp;
                    if (na && nb)
                    {
                        cmp = ia.CompareTo(ib);
                    }
                    else
                    {
                        cmp = string.Compare(ta, tb, StringComparison.InvariantCultureIgnoreCase);
                    }

                    if (cmp != 0) return cmp;
                }

                return ax.Count.CompareTo(bx.Count);
            }

            private static List<string> Tokenize(string s)
            {
                // Splits into sequences of digits and non-digits
                // Example: "A-10+2" => ["A-", "10", "+", "2"]
                var list = new List<string>();
                foreach (Match m in Regex.Matches(s, @"\d+|\D+"))
                    list.Add(m.Value);
                return list;
            }
        }

        #endregion

        #region Config: Export path per project
        private bool SheetsFilter(object obj)
        {
            var item = obj as SheetItem;
            if (item == null) return false;

            var q = (SheetSearchTextBox?.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(q)) return true;

            q = q.ToLowerInvariant();

            // Search fields (safe even if some are null)
            var number = (item.Number ?? string.Empty).ToLowerInvariant();
            var name = (item.Name ?? string.Empty).ToLowerInvariant();
            var version = (item.Version ?? string.Empty).ToLowerInvariant();
            var size = (item.SheetSize ?? string.Empty).ToLowerInvariant();
            var viewName = (item.ViewName ?? string.Empty).ToLowerInvariant();

            return number.Contains(q)
                || name.Contains(q)
                || version.Contains(q)
                || size.Contains(q)
                || viewName.Contains(q);
        }

        private void SheetSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _sheetsView?.Refresh();
            UpdateClearButtonState();
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            SheetSearchTextBox.Text = string.Empty;
            SheetSearchTextBox.Focus();

            _sheetsView?.Refresh();
            UpdateClearButtonState();
        }

        private void UpdateClearButtonState()
        {
            if (ClearSearchButton == null || SheetSearchTextBox == null) return;
            ClearSearchButton.IsEnabled = !string.IsNullOrWhiteSpace(SheetSearchTextBox.Text);
        }

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
            _sheetsToScan.Clear();

            // 1. Get all sheets
            var sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            // 2. Get all viewports in the document to map SheetId -> ViewName
            //    This avoids the N+1 query inside the loop.
            var viewports = new FilteredElementCollector(_doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .ToList();

            // Map: SheetId -> List<ViewId>
            var sheetViews = new Dictionary<ElementId, ElementId>();
            foreach (var vp in viewports)
            {
                if (!sheetViews.ContainsKey(vp.SheetId))
                {
                    sheetViews[vp.SheetId] = vp.ViewId; // store first view only
                }
            }

            foreach (var sheet in sheets)
            {
                string viewName = "";
                if (sheetViews.TryGetValue(sheet.Id, out ElementId viewId))
                {
                    var view = _doc.GetElement(viewId) as View;
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

                // Quick check for A4 merge sheets
                if ((sheet.SheetNumber ?? "").Contains("--"))
                {
                    item.SheetSize = "A4";
                }
                else
                {
                    // Defer calculation
                    item.SheetSize = "Scanning..."; 
                    _sheetsToScan.Enqueue(item);
                }

                // Version / Revision (can also be slow, but usually okay-ish)
                string versionText = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION)?.AsString();
                if (string.IsNullOrWhiteSpace(versionText))
                {
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

        private void OnIdling_ScanPaperSizes(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            if (_sheetsToScan.Count == 0)
            {
                // Done scanning, unsubscribe
                if (_uiapp != null)
                {
                    _uiapp.Idling -= OnIdling_ScanPaperSizes;
                    _uiapp = null;
                }
                return;
            }

            // Process a chunk of sheets
            int count = 0;
            while (count < SheetsPerIdleTick && _sheetsToScan.Count > 0)
            {
                var item = _sheetsToScan.Dequeue();
                var sheet = _doc.GetElement(item.Id) as ViewSheet;
                
                if (sheet != null)
                {
                    // Logic from PaperSizeHelper, now run in background
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
                }
                count++;
            }
            
            // If we have more work, ensure we get called back quickly
            if (_sheetsToScan.Count > 0)
            {
                e.SetRaiseWithoutDelay();
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
            _sheetsView?.Refresh();

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
            var exportDxf = (LogicalTreeHelper.FindLogicalNode(this, "DxfExportCheckbox") as CheckBox)?.IsChecked == true;

            if (!exportDwg && !exportPdf && !exportDxf)
            {
                RevitTaskDialog.Show("Info", "Neither DWG, DXF nor PDF export is selected.");
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
            _exportHandler.ExportDxf = exportDxf;
            _exportHandler.TemplateDxfPath = _templateDxfPath;

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
                
                // Cleanup Idling if still active
                if (_uiapp != null)
                {
                    _uiapp.Idling -= OnIdling_ScanPaperSizes;
                    _uiapp = null;
                }

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
