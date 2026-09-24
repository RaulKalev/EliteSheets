using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using EliteSheets.ExternalEvents;
using EliteSheets.Helpers;
using EliteSheets.Models;
using EliteSheets.Services;
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
                MessageBox.Show(Loc.Format("ThemeLoadError", ex.Message, themeUri), Loc.Get("ThemeLoadErrorTitle"),
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

        private void LanguageToggle_Click(object sender, RoutedEventArgs e)
        {
            // Saved user-wide; bound text in this window updates immediately
            Loc.Save(Loc.Language == Loc.English ? Loc.Estonian : Loc.English);
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

                        if (config.TryGetValue(ConvertMergedToDwgKey, out var convertObj) && convertObj is bool convert)
                            ConvertMergedToDwgCheckbox.IsChecked = convert;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.Format("ConfigLoadError", ex.Message), Loc.Get("LoadErrorTitle"),
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
                MessageBox.Show(Loc.Format("SettingsSaveError", ex.Message), Loc.Get("SaveErrorTitle"),
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


            ThemedMessageDialog.Show(this, Loc.Get("TemplateSetupTitle"), Loc.Get("TemplateSetupMessage"), ThemedDialogKind.Info);

            var dlg = new OpenFileDialog
            {
                Title = Loc.Get("TemplatePickTitle"),
                Filter = Loc.Get("DxfFileFilter"),
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
                RevitTaskDialog.Show(Loc.Get("SaveErrorTitle"), Loc.Format("ExportSettingsSaveError", ex.Message));
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
                RevitTaskDialog.Show(Loc.Get("LoadErrorTitle"), Loc.Format("ExportPathLoadError", ex.Message));
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

        private const string ConvertMergedToDwgKey = "ConvertMergedToDwg";

        private void ConvertMergedToDwgCheckbox_Click(object sender, RoutedEventArgs e)
        {
            if (ConvertMergedToDwgCheckbox.IsChecked == true &&
                EliteSheets.Services.OdaConverterService.FindConverter() == null)
            {
                ConvertMergedToDwgCheckbox.IsChecked = false;
                ThemedMessageDialog.Show(this, Loc.Get("OdaNotFoundTitle"), Loc.Get("OdaNotFoundMessage"));
                return;
            }

            SaveGlobalSetting(ConvertMergedToDwgKey, ConvertMergedToDwgCheckbox.IsChecked == true);
        }

        /// <summary>
        /// Stores a user-wide value at the root of the config, preserving all other keys.
        /// </summary>
        private void SaveGlobalSetting(string key, object value)
        {
            try
            {
                var config = new Dictionary<string, object>();
                if (File.Exists(ConfigFilePath))
                {
                    config = JsonConvert.DeserializeObject<Dictionary<string, object>>(File.ReadAllText(ConfigFilePath))
                             ?? new Dictionary<string, object>();
                }

                config[key] = value;

                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFilePath));
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show(Loc.Get("SaveErrorTitle"), Loc.Format("SettingsSaveError", ex.Message));
            }
        }

        #endregion

        #region Data loading

        private void LoadSheets()
        {
            Sheets.Clear();
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
                RevitTaskDialog.Show(Loc.Get("ErrorTitle"), Loc.Format("DwgSetupsLoadError", ex.Message));
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
                Autodesk.Revit.UI.TaskDialog.Show(Loc.Get("ReloadErrorTitle"), Loc.Format("ReloadError", ex.Message));
            }
            _sheetsView?.Refresh();

        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder(Loc.Get("BrowseExportFolderTooltip"), ExportPathTextBox, ExportPathsKey);

        private void BrowsePdfButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder(Loc.Get("BrowsePdfFolderTooltip"), PdfExportPathTextBox, PdfExportPathsKey);

        private void BrowseDwgButton_Click(object sender, RoutedEventArgs e)
            => BrowseForFolder(Loc.Get("BrowseDwgFolderTooltip"), DwgExportPathTextBox, DwgExportPathsKey);

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
                RevitTaskDialog.Show(Loc.Get("BrowseErrorTitle"), ex.Message);
            }
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            // 1) validate selection + export options
            var selected = Sheets.Where(s => s.IsChecked).ToList();
            if (!selected.Any())
            {
                ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("NoSheetsSelected"));
                return;
            }

            var exportDwg = (LogicalTreeHelper.FindLogicalNode(this, "DwgExportCheckbox") as CheckBox)?.IsChecked == true;
            var exportPdf = (LogicalTreeHelper.FindLogicalNode(this, "PdfExportCheckbox") as CheckBox)?.IsChecked == true;
            var exportDxf = false; // DXF button removed

            if (!exportDwg && !exportPdf && !exportDxf)
            {
                ThemedMessageDialog.Show(this, Loc.Get("InfoTitle"), Loc.Get("NoFormatSelected"), ThemedDialogKind.Info);
                return;
            }

            string exportPath = null, pdfExportPath = null, dwgExportPath = null;
            if (UseSeparateFolders)
            {
                pdfExportPath = PdfExportPathTextBox.Text?.Trim();
                dwgExportPath = DwgExportPathTextBox.Text?.Trim();

                if (exportPdf && (string.IsNullOrWhiteSpace(pdfExportPath) || !Directory.Exists(pdfExportPath)))
                {
                    ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("InvalidPdfFolder"));
                    return;
                }
                if ((exportDwg || exportDxf) && (string.IsNullOrWhiteSpace(dwgExportPath) || !Directory.Exists(dwgExportPath)))
                {
                    ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("InvalidDwgFolder"));
                    return;
                }
            }
            else
            {
                exportPath = ExportPathTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(exportPath) || !Directory.Exists(exportPath))
                {
                    ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("InvalidExportFolder"));
                    return;
                }
            }

            var exportSetup = DwgExportComboBox.SelectedItem as ExportDWGSettings;
            if (exportDwg && exportSetup == null)
            {
                ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("NoDwgSetup"));
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
                var message = Loc.Get("InvalidNumbersIntro") + "\n\n" +
                              string.Join("\n", invalidSheets.Select(i => $"• \"{i.Sheet.Number}\" → {i.Invalid}")) +
                              "\n\n" + Loc.Get("InvalidNumbersWindows") + "\n" +
                              string.Join(" ", forbidden.Select(c => $"'{c}'")) +
                              "\n\n" + Loc.Get("InvalidNumbersQuestion");

                if (!ThemedMessageDialog.AskYesNo(this, Loc.Get("InvalidNumbersTitle"), message, ThemedDialogKind.Warning))
                    return;

                selected = selected.Except(invalidSheets.Select(i => i.Sheet)).ToList();
                if (!selected.Any())
                {
                    ThemedMessageDialog.Show(this, Loc.Get("NoValidSheetsTitle"), Loc.Get("NoValidSheets"));
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
                ThemedMessageDialog.Show(this, Loc.Get("ErrorTitle"), Loc.Get("SheetsNotResolved"), ThemedDialogKind.Error);
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
            _exportHandler.ConvertMergedToDwg = ConvertMergedToDwgCheckbox.IsChecked == true;
            _exportHandler.OwnerWindow = this;

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

            if (newState)
            {
                // If checking, only check visible items
                if (_sheetsView != null)
                {
                    foreach (var item in _sheetsView)
                    {
                        var sheetItem = item as SheetItem;
                        if (sheetItem != null)
                        {
                            sheetItem.IsChecked = true;
                        }
                    }
                }
            }
            else
            {
                // If unchecking, uncheck all items globally
                foreach (var item in Sheets)
                {
                    item.IsChecked = false;
                }
            }

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
