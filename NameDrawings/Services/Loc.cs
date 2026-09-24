using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace EliteSheets.Services
{
    /// <summary>
    /// UI string lookup for English ("en") and Estonian ("et").
    /// The chosen language is stored user-wide as "Language" in config.json.
    /// </summary>
    public static class Loc
    {
        public const string English = "en";
        public const string Estonian = "et";

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\EliteSheets\config.json";
        private const string LanguageKey = "Language";

        /// <summary>Current language. Estonian unless the user switched to English.</summary>
        public static string Language { get; private set; } = Estonian;

        /// <summary>Raised after <see cref="Save"/> switches the language.</summary>
        public static event EventHandler LanguageChanged;

        private static readonly Dictionary<string, (string En, string Et)> Strings = new Dictionary<string, (string, string)>
        {
            // Window / title bar / ribbon
            ["AppTitle"] = ("Sheet printing", "Jooniste printimine"),
            ["RibbonText"] = ("Sheet\nPrinting", "Jooniste\nPrintimine"),
            ["RibbonTooltip"] = ("Sheet printing.", "Jooniste printimine."),
            ["CloseWindow"] = ("Close window", "Sulge aken"),
            ["Close"] = ("Close", "Sulge"),
            ["MinimizeWindow"] = ("Minimize window", "Minimeeri aken"),
            ["Minimize"] = ("Minimize", "Minimeeri"),

            // Top bar
            ["SearchPlaceholder"] = ("Search...", "Otsi..."),
            ["SearchSheets"] = ("Search sheets", "Otsi jooniseid"),
            ["SearchTooltip"] = ("Search (Number / Name / Version / Size)", "Otsi (number / nimi / versioon / suurus)"),
            ["ClearSearch"] = ("Clear search", "Tühjenda otsing"),
            ["Clear"] = ("Clear", "Tühjenda"),
            ["ReloadSheets"] = ("Reload sheets", "Laadi joonised uuesti"),
            ["Reload"] = ("Reload", "Laadi uuesti"),
            ["ToggleTheme"] = ("Toggle theme", "Vaheta teemat"),
            ["LanguageShort"] = ("EN", "ET"),
            ["SwitchLanguage"] = ("Switch language to Estonian", "Vaheta keel inglise keeleks"),

            // Sheet grid
            ["SelectAllSheets"] = ("Select all sheets", "Vali kõik joonised"),
            ["CheckUncheckAll"] = ("Check/Uncheck all", "Märgi/eemalda kõik"),
            ["SelectSheetFormat"] = ("Select sheet {0} - {1}", "Vali joonis {0} - {1}"),
            ["ColSheetNumber"] = ("Sheet number", "Joonise number"),
            ["ColSheetName"] = ("Sheet name", "Joonise nimi"),
            ["ColVersion"] = ("Version", "Versioon"),

            // Export controls
            ["ExportToDwg"] = ("Export to DWG", "Ekspordi DWG-na"),
            ["ExportToDwgTooltip"] = ("Export to DWG format", "Ekspordi DWG formaadis"),
            ["ExportToPdf"] = ("Export to PDF", "Ekspordi PDF-na"),
            ["ExportToPdfTooltip"] = ("Export to PDF format", "Ekspordi PDF formaadis"),
            ["DwgExportSettings"] = ("DWG Export settings:", "DWG ekspordi seaded:"),
            ["SelectDwgSetup"] = ("Select DWG export setup", "Vali DWG ekspordi seadistus"),
            ["Browse"] = ("Browse", "Sirvi"),
            ["BrowseExportFolder"] = ("Browse for export folder", "Vali ekspordi kaust"),
            ["BrowseExportFolderTooltip"] = ("Select a folder for exported files", "Vali kaust eksporditud failidele"),
            ["ExportPath"] = ("Export path", "Ekspordi asukoht"),
            ["ExportPathTooltip"] = ("Selected export path", "Valitud ekspordi asukoht"),
            ["BrowsePdfFolder"] = ("Browse for PDF export folder", "Vali PDF-ide kaust"),
            ["BrowsePdfFolderTooltip"] = ("Select a folder for exported PDF files", "Vali kaust eksporditud PDF-failidele"),
            ["PdfExportPath"] = ("PDF export path", "PDF-ide asukoht"),
            ["PdfExportPathTooltip"] = ("Selected PDF export path", "Valitud PDF-ide asukoht"),
            ["BrowseDwgFolder"] = ("Browse for DWG export folder", "Vali DWG-de kaust"),
            ["BrowseDwgFolderTooltip"] = ("Select a folder for exported DWG files", "Vali kaust eksporditud DWG-failidele"),
            ["DwgExportPath"] = ("DWG export path", "DWG-de asukoht"),
            ["DwgExportPathTooltip"] = ("Selected DWG export path", "Valitud DWG-de asukoht"),
            ["SeparateFolders"] = ("Separate PDF / DWG folders", "Eraldi PDF / DWG kaustad"),
            ["SeparateFoldersName"] = ("Use separate folders for PDF and DWG", "Kasuta PDF-ide ja DWG-de jaoks eraldi kaustu"),
            ["SeparateFoldersTooltip"] = ("Export PDF and DWG files to different folders", "Ekspordi PDF- ja DWG-failid erinevatesse kaustadesse"),
            ["MergedAsDwg"] = ("Merged drawings as DWG (ODA)", "Liidetud joonised DWG-na (ODA)"),
            ["MergedAsDwgName"] = ("Convert merged drawings to DWG with ODA File Converter", "Teisenda liidetud joonised ODA File Converteriga DWG-ks"),
            ["MergedAsDwgTooltip"] = ("Convert merged drawings from DXF to DWG (AutoCAD 2018) using ODA File Converter",
                                      "Teisenda liidetud joonised DXF-ist DWG-ks (AutoCAD 2018) ODA File Converteriga"),
            ["Print"] = ("Print", "Prindi"),
            ["PrintTooltip"] = ("Print and export selected sheets", "Prindi ja ekspordi valitud joonised"),

            // Window chrome, toolbar, list and action bar
            ["Maximize"] = ("Maximize", "Maksimeeri"),
            ["Restore"] = ("Restore down", "Taasta"),
            ["ThemeDarkTooltip"] = ("Dark appearance – switch to light", "Tume välimus – lülita heledale"),
            ["ThemeLightTooltip"] = ("Light appearance – switch to dark", "Hele välimus – lülita tumedale"),
            ["SearchHint"] = ("Search number, name or version", "Otsi numbri, nime või versiooni järgi"),
            ["SearchTooltipShortcut"] = ("Search sheets by number, name, version or view (Ctrl+F)", "Otsi jooniseid numbri, nime, versiooni või vaate järgi (Ctrl+F)"),
            ["ClearSearchShortcut"] = ("Clear search (Esc)", "Tühjenda otsing (Esc)"),
            ["ReloadTooltip"] = ("Reload sheets and DWG setups from the model", "Laadi joonised ja DWG seadistused mudelist uuesti"),
            ["SelectAllVisible"] = ("Select or clear all visible sheets", "Vali või eemalda kõik nähtavad joonised"),
            ["NoSheetsInModel"] = ("No sheets in this model", "Mudelis pole jooniseid"),
            ["ChooseFormat"] = ("Choose PDF or DWG to export", "Vali eksportimiseks PDF või DWG"),
            ["NoneSelected"] = ("{0} sheets · none selected", "{0} joonist · ühtegi pole valitud"),
            ["SelectedOfTotal"] = ("{0} of {1} selected", "{0} / {1} valitud"),
            ["PrintOne"] = ("Print 1 sheet", "Prindi 1 joonis"),
            ["PrintMany"] = ("Print {0} sheets", "Prindi {0} joonist"),
            ["PrintShortcutTooltip"] = ("Print and export the selected sheets (Ctrl+Enter)", "Prindi ja ekspordi valitud joonised (Ctrl+Enter)"),
            ["EmptySearchTitle"] = ("No matching sheets", "Sobivaid jooniseid ei leitud"),
            ["EmptySearchDetail"] = ("Nothing matches “{0}”. Try a sheet number, part of a name or a version.",
                                     "„{0}“ ei leitud. Proovi joonise numbrit, osa nimest või versiooni."),
            ["EmptyModelTitle"] = ("No sheets", "Jooniseid pole"),
            ["EmptyModelDetail"] = ("This model has no sheets yet. Create sheets in Revit, then reload.",
                                    "Selles mudelis pole veel jooniseid. Loo joonised Revitis ja laadi uuesti."),
            ["FormatLabel"] = ("Format", "Formaat"),
            ["DwgSetupLabel"] = ("DWG setup", "DWG seadistus"),
            ["DwgSetupTooltip"] = ("DWG export setup defined in the Revit project", "Revit projektis määratud DWG ekspordi seadistus"),
            ["DestinationLabel"] = ("Destination", "Sihtkoht"),
            ["SeparateFoldersShort"] = ("Separate folders", "Eraldi kaustad"),
            ["Choose"] = ("Choose…", "Vali…"),
            ["NoFolderChosen"] = ("No folder chosen", "Kausta pole valitud"),
            ["ExportPdfChip"] = ("Export selected sheets as PDF", "Ekspordi valitud joonised PDF-na"),
            ["ExportDwgChip"] = ("Export selected sheets as DWG", "Ekspordi valitud joonised DWG-na"),

            // What's new (changelog popup)
            ["WhatsNewTitle"] = ("What's new", "Mis on uut"),
            ["WhatsNewVersion"] = ("Version {0} · {1}", "Versioon {0} · {1}"),
            ["WhatsNewOlder"] = ("Version {0}", "Versioon {0}"),
            ["GotIt"] = ("Got it", "Selge"),

            // Dialog buttons
            ["Yes"] = ("Yes", "Jah"),
            ["No"] = ("No", "Ei"),
            ["Ok"] = ("OK", "OK"),
            ["CloseDialog"] = ("Close dialog", "Sulge aken"),

            // ODA
            ["OdaNotFoundTitle"] = ("ODA File Converter not found", "ODA File Converterit ei leitud"),
            ["OdaNotFoundMessage"] = (
                "ODA File Converter is not installed on this computer, so merged drawings can't be converted to DWG.\n\n" +
                "Install it from opendesign.com (default location: C:\\Program Files\\ODA) and try again.",
                "ODA File Converter ei ole selles arvutis paigaldatud, seega liidetud jooniseid ei saa DWG-ks teisendada.\n\n" +
                "Paigalda see aadressilt opendesign.com (vaikimisi asukoht: C:\\Program Files\\ODA) ja proovi uuesti."),

            // Template setup
            ["TemplateSetupTitle"] = ("Setup required", "Vajalik seadistus"),
            ["TemplateSetupMessage"] = (
                "Please choose the DXF template KilbiTemplate.dxf (must be .dxf). The suggested location opens in the dialog automatically.\n" +
                "The file is in the folder:  \\EULE Dropbox\\0_EULE  Team folder (kogu kollektiiv)\\02_EULE REVIT TEMPLATE ",
                "Palun vali DXF template KilbiTemplate.dxf (tingimata .dxf). Soovitatav asukoht avatakse dialoogis automaatselt.\n" +
                "Fail asub kaustas:  \\EULE Dropbox\\0_EULE  Team folder (kogu kollektiiv)\\02_EULE REVIT TEMPLATE "),
            ["TemplatePickTitle"] = ("Choose template (DXF)", "Vali mall (DXF)"),
            ["DxfFileFilter"] = ("DXF file (*.dxf)|*.dxf", "DXF fail (*.dxf)|*.dxf"),

            // Generic errors
            ["ThemeLoadErrorTitle"] = ("Theme load error", "Teema laadimise viga"),
            ["ThemeLoadError"] = ("Failed to load theme: {0}", "Teema laadimine ebaõnnestus: {0}"),
            ["LoadErrorTitle"] = ("Load error", "Laadimise viga"),
            ["ConfigLoadError"] = ("Failed to load theme/config: {0}", "Teema/seadete laadimine ebaõnnestus: {0}"),
            ["ExportPathLoadError"] = ("Failed to load export path:\n{0}", "Ekspordi asukoha laadimine ebaõnnestus:\n{0}"),
            ["SaveErrorTitle"] = ("Save error", "Salvestamise viga"),
            ["SettingsSaveError"] = ("Failed to save settings:\n{0}", "Seadete salvestamine ebaõnnestus:\n{0}"),
            ["ExportSettingsSaveError"] = ("Failed to save export settings:\n{0}", "Ekspordi seadete salvestamine ebaõnnestus:\n{0}"),
            ["ErrorTitle"] = ("Error", "Viga"),
            ["InfoTitle"] = ("Info", "Info"),
            ["DwgSetupsLoadError"] = ("Failed to load DWG export setups: {0}", "DWG ekspordi seadistuste laadimine ebaõnnestus: {0}"),
            ["ReloadErrorTitle"] = ("Reload error", "Uuesti laadimise viga"),
            ["ReloadError"] = ("Failed to reload data: {0}", "Andmete uuesti laadimine ebaõnnestus: {0}"),
            ["BrowseErrorTitle"] = ("Browse error", "Kausta valimise viga"),

            // Export validation
            ["NoSheetsSelected"] = ("No sheets selected.", "Ühtegi joonist pole valitud."),
            ["NoFormatSelected"] = ("Neither DWG, DXF nor PDF export is selected.", "Ükski eksporditav formaat (DWG, DXF, PDF) pole valitud."),
            ["InvalidPdfFolder"] = ("Please select a valid PDF export folder.", "Palun vali kehtiv PDF-ide kaust."),
            ["InvalidDwgFolder"] = ("Please select a valid DWG export folder.", "Palun vali kehtiv DWG-de kaust."),
            ["InvalidExportFolder"] = ("Please select a valid export folder.", "Palun vali kehtiv ekspordi kaust."),
            ["NoDwgSetup"] = ("Please select a DWG export setup.", "Palun vali DWG ekspordi seadistus."),
            ["InvalidNumbersTitle"] = ("Invalid sheet numbers", "Mittesobivad joonise numbrid"),
            ["InvalidNumbersIntro"] = ("The following sheets have invalid characters in their sheet number:", "Järgnevatel lehtedel on mittesobivad märgid nende lehenumbris:"),
            ["InvalidNumbersWindows"] = ("Windows does not allow these characters in file names:", "Windows ei luba järgmisi märke failinimedes:"),
            ["InvalidNumbersQuestion"] = ("Do you want to export the remaining sheets?", "Kas soovid eksportida ülejäänud lehed?"),
            ["NoValidSheetsTitle"] = ("No valid sheets", "Sobivaid jooniseid pole"),
            ["NoValidSheets"] = ("All selected sheets have invalid characters. Nothing to export.", "Kõigil valitud lehtedel on mittesobivad märgid. Midagi ei eksporditud."),
            ["SheetsNotResolved"] = ("Could not resolve selected sheets in the document.", "Valitud jooniseid ei leitud dokumendist."),

            // Export results
            ["ExportDoneTitle"] = ("Export finished", "Eksport lõppenud"),
            ["ExportDoneMessage"] = ("Export finished.\n\nOpen the export folder?", "Eksport lõppenud.\n\nAvada ekspordi kaust?"),
            ["ExportFailedTitle"] = ("EliteSheets - Export failed", "EliteSheets - eksport ebaõnnestus"),
            ["ExportFailed"] = ("Export failed for all selected sheets.", "Eksport ebaõnnestus kõigi valitud jooniste puhul."),
            ["DwgExportErrorsTitle"] = ("DWG export errors", "DWG ekspordi vead"),
            ["DxfMergeErrorsTitle"] = ("DXF merge errors", "DXF liitmise vead"),
            ["SheetError"] = ("Sheet {0}: {1}", "Joonis {0}: {1}"),
            ["DxfMergeTitle"] = ("DXF export merge", "DXF eksport merge"),
            ["TemplateMissing"] = ("The template file (DXF) location is not set or the file was not found. Open the EliteSheets window and choose the template in settings.",
                                   "Mallifaili (DXF) asukoht ei ole seadistatud või faili ei leitud. Ava EliteSheets aken ja vali mall seadetes."),
            ["OdaMissingSavedDxf"] = ("ODA File Converter not found - merged drawings were saved as DXF.", "ODA File Converterit ei leitud - liidetud joonised salvestati DXF-na."),
            ["DxfNotFoundForMerge"] = ("DXF for sheet '{0}' not found for merging.", "Joonise '{0}' DXF-i ei leitud liitmiseks."),
            ["MergeFailed"] = ("Merge failed for group {0}: {1}", "Grupi {0} liitmine ebaõnnestus: {1}"),
            ["DwgConversionFailed"] = ("DWG conversion failed: {0}", "DWG-ks teisendamine ebaõnnestus: {0}"),
            ["NotConvertedSavedDxf"] = ("{0} could not be converted to DWG - saved as DXF.", "{0} ei õnnestunud DWG-ks teisendada - salvestati DXF-na."),
            ["PrintErrorTitle"] = ("Print error", "Printimise viga"),
            ["PrintSettingsError"] = ("Failed to apply print settings:\n{0}", "Prindiseadete rakendamine ebaõnnestus:\n{0}"),
        };

        public static string Get(string key)
        {
            if (!Strings.TryGetValue(key, out var s)) return key;
            return Language == English ? s.En : s.Et;
        }

        public static string Format(string key, params object[] args) => string.Format(Get(key), args);

        /// <summary>Reads the saved language from config.json.</summary>
        public static void Load()
        {
            try
            {
                if (!File.Exists(ConfigFilePath)) return;
                var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(File.ReadAllText(ConfigFilePath));
                if (config != null && config.TryGetValue(LanguageKey, out var raw) &&
                    (raw as string == English || raw as string == Estonian))
                {
                    Language = (string)raw;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to load language setting.", ex);
            }
        }

        /// <summary>Sets and saves the language user-wide, preserving other config keys.</summary>
        public static void Save(string language)
        {
            Language = language == English ? English : Estonian;
            LanguageChanged?.Invoke(null, EventArgs.Empty);
            try
            {
                var config = new Dictionary<string, object>();
                if (File.Exists(ConfigFilePath))
                {
                    config = JsonConvert.DeserializeObject<Dictionary<string, object>>(File.ReadAllText(ConfigFilePath))
                             ?? new Dictionary<string, object>();
                }

                config[LanguageKey] = Language;

                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFilePath));
                File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to save language setting.", ex);
            }
        }
    }
}
