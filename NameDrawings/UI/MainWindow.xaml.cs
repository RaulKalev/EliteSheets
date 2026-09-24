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

        private bool _isDarkMode = true;
        private CheckBox _checkAllBox;
        private string _templateDxfPath = string.Empty;

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
            // Palette (slot 0, swapped by LoadTheme) and shared styles must exist before the XAML's StaticResources.
            Resources.MergedDictionaries.Add(Themes.ThemeResources.Palette(dark: true));
            Resources.MergedDictionaries.Add(Themes.ThemeResources.Styles());
            InitializeComponent();

            _uiDoc = uiDoc;
            _doc = doc;
            _currentView = currentView;

            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // Window infrastructure (resize, snap and drag come from WindowChrome)
            Closed += MainWindow_Closed;
            StateChanged += (s, e) => UpdateMaximizedState();
            PreviewKeyDown += MainWindow_PreviewKeyDown;
            ProjectText.Text = doc.Title;

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
            UpdateSelectionSummary();

            LoadDwgExportSetups();


            // External events
            _exportHandler = new ExportSheetsHandler();
            _exportEvent = ExternalEvent.Create(_exportHandler);

            _createPrintSettingHandler = new CreatePrintSettingHandler { Doc = _doc };
            _createPrintSettingEvent = ExternalEvent.Create(_createPrintSettingHandler);

        }

        #endregion

        #region Theme

        /// <summary>
        /// Swaps only the palette dictionary; the shared styles stay merged and pick up the new colours through
        /// DynamicResource.
        /// </summary>
        private void LoadTheme()
        {
            try
            {
                var palette = Themes.ThemeResources.Palette(_isDarkMode);
                var merged = Resources.MergedDictionaries;
                var existing = merged.FirstOrDefault(Themes.ThemeResources.IsPalette);
                if (existing != null) merged[merged.IndexOf(existing)] = palette;
                else merged.Insert(0, palette);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load theme: {ex.Message}", "Theme Load Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            ThemeIcon.Kind = _isDarkMode
                ? MaterialDesignThemes.Wpf.PackIconKind.WeatherNight
                : MaterialDesignThemes.Wpf.PackIconKind.WhiteBalanceSunny;
            ThemeButton.ToolTip = _isDarkMode ? "Dark appearance \u2013 switch to light" : "Light appearance \u2013 switch to dark";
        }

        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = !_isDarkMode;
            LoadTheme();
            SaveThemeState();
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
            var viewName = (item.ViewName ?? string.Empty).ToLowerInvariant();
            return number.Contains(q)
                || name.Contains(q)
                || version.Contains(q)
                || viewName.Contains(q);
        }

        private void SheetSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _sheetsView?.Refresh();
            UpdateClearButtonState();
            UpdateSelectionSummary();
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            SheetSearchTextBox.Text = string.Empty;
            SheetSearchTextBox.Focus();

            _sheetsView?.Refresh();
            UpdateClearButtonState();
            UpdateSelectionSummary();
        }

        private void UpdateClearButtonState()
        {
            if (ClearSearchButton == null || SheetSearchTextBox == null) return;
            ClearSearchButton.IsEnabled = !string.IsNullOrWhiteSpace(SheetSearchTextBox.Text);
        }

        private const string ExportPathsKey = "ExportPaths";
        private const string PdfExportPathsKey = "PdfExportPaths";
        private const string DwgExportPathsKey = "DwgExportPaths";
        private const string SeparateExportFoldersKey = "SeparateExportFolders";

        /// <summary>
        /// Stores a value under config[sectionKey][projectName], preserving all other keys.
        /// </summary>
        private void SaveProjectSetting(string sectionKey, string value)
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

                Dictionary<string, string> section;
                if (config.TryGetValue(sectionKey, out object rawSection) &&
                    rawSection is Newtonsoft.Json.Linq.JObject jObj)
                {
                    section = jObj.ToObject<Dictionary<string, string>>();
                }
                else
                {
                    section = new Dictionary<string, string>();
                }

                section[projectName] = value;
                config[sectionKey] = section;

                if (!config.ContainsKey("IsDarkMode"))
                    config["IsDarkMode"] = _isDarkMode;

                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFilePath));
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Save Error", $"Failed to save export settings:\n{ex.Message}");
            }
        }

        private void LoadExportPathForCurrentProject()
        {
            try
            {
                if (!File.Exists(ConfigFilePath)) return;

                var json = File.ReadAllText(ConfigFilePath);
                var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (config == null) return;

                var projectName = Path.GetFileName(_doc.PathName);
                if (string.IsNullOrEmpty(projectName)) return;

                string Get(string sectionKey)
                {
                    if (config.TryGetValue(sectionKey, out var raw) &&
                        raw is Newtonsoft.Json.Linq.JObject jObj &&
                        jObj.ToObject<Dictionary<string, string>>().TryGetValue(projectName, out var v))
                        return v;
                    return null;
                }

                var savedPath = Get(ExportPathsKey);
                if (savedPath != null) ExportPathTextBox.Text = savedPath;

                PdfExportPathTextBox.Text = Get(PdfExportPathsKey) ?? string.Empty;
                DwgExportPathTextBox.Text = Get(DwgExportPathsKey) ?? string.Empty;

                bool separate = string.Equals(Get(SeparateExportFoldersKey), "true", StringComparison.OrdinalIgnoreCase);
                SeparateFoldersCheckbox.IsChecked = separate;
                UpdateFolderPanels();
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("Load Error", $"Failed to load export path:\n{ex.Message}");
            }
        }

        private bool UseSeparateFolders => SeparateFoldersCheckbox?.IsChecked == true;

        private void UpdateFolderPanels()
        {
            SingleFolderPanel.Visibility = UseSeparateFolders ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
            SeparateFoldersPanel.Visibility = UseSeparateFolders ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
        }

        private void SeparateFoldersCheckbox_Click(object sender, RoutedEventArgs e)
        {
            // Seed empty paths from the shared folder so switching modes is painless
            if (UseSeparateFolders)
            {
                var shared = ExportPathTextBox.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(shared) && Directory.Exists(shared))
                {
                    if (string.IsNullOrWhiteSpace(PdfExportPathTextBox.Text))
                    {
                        PdfExportPathTextBox.Text = shared;
                        SaveProjectSetting(PdfExportPathsKey, shared);
                    }
                    if (string.IsNullOrWhiteSpace(DwgExportPathTextBox.Text))
                    {
                        DwgExportPathTextBox.Text = shared;
                        SaveProjectSetting(DwgExportPathsKey, shared);
                    }
                }
            }

            UpdateFolderPanels();
            SaveProjectSetting(SeparateExportFoldersKey, UseSeparateFolders ? "true" : "false");
        }

        #endregion

        #region Data loading

        private void LoadSheets()
        {
            foreach (var old in Sheets) old.PropertyChanged -= SheetItem_PropertyChanged;
            Sheets.Clear();

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

                item.PropertyChanged += SheetItem_PropertyChanged;
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
            _sheetsView?.Refresh();
            UpdateSelectionSummary();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder("Select export folder", ExportPathTextBox, ExportPathsKey);

        private void BrowsePdfButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder("Select PDF export folder", PdfExportPathTextBox, PdfExportPathsKey);

        private void BrowseDwgButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder("Select DWG export folder", DwgExportPathTextBox, DwgExportPathsKey);

        private void BrowseForFolder(string description, System.Windows.Controls.TextBox target, string configKey)
        {
            try
            {
                using (var dlg = new WinForms.FolderBrowserDialog())
                {
                    dlg.Description = description;
                    dlg.ShowNewFolderButton = true;

                    var initial = target.Text;
                    if (string.IsNullOrWhiteSpace(initial) || !Directory.Exists(initial))
                        initial = ExportPathTextBox.Text;
                    dlg.SelectedPath = (!string.IsNullOrWhiteSpace(initial) && Directory.Exists(initial))
                        ? initial
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    if (dlg.ShowDialog() == WinForms.DialogResult.OK)
                    {
                        target.Text = dlg.SelectedPath;
                        SaveProjectSetting(configKey, dlg.SelectedPath);
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

            var exportDwg = (LogicalTreeHelper.FindLogicalNode(this, "DwgExportCheckbox") as CheckBox)?.IsChecked == true;
            var exportPdf = (LogicalTreeHelper.FindLogicalNode(this, "PdfExportCheckbox") as CheckBox)?.IsChecked == true;
            var exportDxf = false; // DXF button removed

            if (!exportDwg && !exportPdf && !exportDxf)
            {
                RevitTaskDialog.Show("Info", "Neither DWG, DXF nor PDF export is selected.");
                return;
            }

            string exportPath = null, pdfExportPath = null, dwgExportPath = null;
            if (UseSeparateFolders)
            {
                pdfExportPath = PdfExportPathTextBox.Text?.Trim();
                dwgExportPath = DwgExportPathTextBox.Text?.Trim();

                if (exportPdf && (string.IsNullOrWhiteSpace(pdfExportPath) || !Directory.Exists(pdfExportPath)))
                {
                    RevitTaskDialog.Show("Error", "Please select a valid PDF export folder.");
                    return;
                }
                if ((exportDwg || exportDxf) && (string.IsNullOrWhiteSpace(dwgExportPath) || !Directory.Exists(dwgExportPath)))
                {
                    RevitTaskDialog.Show("Error", "Please select a valid DWG export folder.");
                    return;
                }
            }
            else
            {
                exportPath = ExportPathTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(exportPath) || !Directory.Exists(exportPath))
                {
                    RevitTaskDialog.Show("Error", "Please select a valid export folder.");
                    return;
                }
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

            // 4) compute paper sizes for selected sheets only
            foreach (var vs in sheetElements)
            {
                var matchingItem = selected.FirstOrDefault(s => s.Id == vs.Id);
                if (matchingItem != null)
                    matchingItem.PaperSize = PaperSizeHelper.GetPaperSizeLabel(vs);
            }

            // 5) raise export event
            _exportHandler.UiDoc = _uiDoc;
            _exportHandler.Doc = _doc;
            _exportHandler.SheetsToExport = sheetElements;
            _exportHandler.ExportPath = exportPath;
            _exportHandler.PdfExportPath = pdfExportPath;
            _exportHandler.DwgExportPath = dwgExportPath;
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
        }

        private void CheckAllBox_Loaded(object sender, RoutedEventArgs e)
        {
            _checkAllBox = sender as CheckBox;
            UpdateSelectionSummary();
        }

        private void CheckAllBox_Click(object sender, RoutedEventArgs e)
        {
            // Decide from the data, not the box: a mixed or empty state selects everything visible, a full one clears.
            var visible = VisibleSheets().ToList();
            var newState = !(visible.Count > 0 && visible.All(s => s.IsChecked));

            if (newState)
            {
                foreach (var item in visible) item.IsChecked = true;
            }
            else
            {
                // Clearing is global, so nothing hidden by the search stays selected by surprise.
                foreach (var item in Sheets) item.IsChecked = false;
            }

            UpdateSelectionSummary();
        }

        private void FormatCheckbox_Click(object sender, RoutedEventArgs e) => UpdateSelectionSummary();

        private void SheetItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SheetItem.IsChecked)) UpdateSelectionSummary();
        }

        private IEnumerable<SheetItem> VisibleSheets() =>
            _sheetsView == null ? Enumerable.Empty<SheetItem>() : _sheetsView.Cast<SheetItem>();

        /// <summary>
        /// Keeps the selection count, the Print button label/state, the select-all box (on / mixed / off) and the
        /// empty state in step with the data. Feedback is inline: the button says what it will do.
        /// </summary>
        private void UpdateSelectionSummary()
        {
            if (SelectionText == null || PrintButton == null) return;

            var selected = Sheets.Count(s => s.IsChecked);
            var anyFormat = PdfExportCheckbox.IsChecked == true || DwgExportCheckbox.IsChecked == true;

            if (Sheets.Count == 0) SelectionText.Text = "No sheets in this model";
            else if (!anyFormat) SelectionText.Text = "Choose PDF or DWG to export";
            else if (selected == 0) SelectionText.Text = $"{Sheets.Count} sheets \u00b7 none selected";
            else SelectionText.Text = $"{selected} of {Sheets.Count} selected";

            PrintButtonText.Text = selected == 0 ? "Print" : selected == 1 ? "Print 1 sheet" : $"Print {selected} sheets";
            PrintButton.IsEnabled = selected > 0 && anyFormat;

            var visible = VisibleSheets().ToList();
            if (_checkAllBox != null)
            {
                var visibleChecked = visible.Count(s => s.IsChecked);
                _checkAllBox.IsChecked = visibleChecked == 0 ? false : visibleChecked == visible.Count ? (bool?)true : null;
            }

            var query = SheetSearchTextBox?.Text?.Trim() ?? string.Empty;
            var isEmpty = visible.Count == 0;
            EmptyState.Visibility = isEmpty ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            if (isEmpty)
            {
                var searching = query.Length > 0;
                EmptyTitle.Text = searching ? "No matching sheets" : "No sheets";
                EmptyDetail.Text = searching
                    ? $"Nothing matches \u201c{query}\u201d. Try a sheet number, part of a name or a version."
                    : "This model has no sheets yet. Create sheets in Revit, then reload.";
                EmptyClearButton.Visibility = searching ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            }
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

        #region Window chrome / keyboard

        private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

        private void Maximize_Click(object sender, RoutedEventArgs e) =>
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void UpdateMaximizedState()
        {
            // A maximized WindowChrome window extends past the work area by the resize frame; pad the content back in.
            var maximized = WindowState == WindowState.Maximized;
            var frame = SystemParameters.WindowResizeBorderThickness;
            RootGrid.Margin = maximized ? new Thickness(frame.Left + 4, frame.Top + 4, frame.Right + 4, frame.Bottom + 4) : new Thickness(0);
            MaximizeIcon.Kind = maximized ? MaterialDesignThemes.Wpf.PackIconKind.WindowRestore : MaterialDesignThemes.Wpf.PackIconKind.WindowMaximize;
            MaximizeButton.ToolTip = maximized ? "Restore down" : "Maximize";
            System.Windows.Automation.AutomationProperties.SetName(MaximizeButton, maximized ? "Restore down" : "Maximize");
        }

        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;

            if (ctrl && e.Key == Key.F)
            {
                SheetSearchTextBox.Focus();
                SheetSearchTextBox.SelectAll();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape && SheetSearchTextBox.IsKeyboardFocusWithin && SheetSearchTextBox.Text.Length > 0)
            {
                ClearSearchButton_Click(sender, e);
                e.Handled = true;
            }
            else if (ctrl && e.Key == Key.Enter && PrintButton.IsEnabled)
            {
                PrintButton_Click(sender, e);
                e.Handled = true;
            }
        }

        /// <summary>Space toggles the checkbox of every selected row (the keyboard equivalent of clicking the box).</summary>
        private void SheetsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Space || e.OriginalSource is CheckBox) return;

            var rows = SheetsDataGrid.SelectedItems.OfType<SheetItem>().ToList();
            if (rows.Count == 0) return;

            var newState = !rows.All(r => r.IsChecked);
            foreach (var row in rows) row.IsChecked = newState;
            e.Handled = true;
        }

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
                foreach (var item in Sheets) item.PropertyChanged -= SheetItem_PropertyChanged;

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
            }
            catch
            {
                // swallow – app is closing
            }
        }

        #endregion
    }
}
