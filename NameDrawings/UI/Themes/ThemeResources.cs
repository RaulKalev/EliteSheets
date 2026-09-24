using System.IO;
using System.Windows;
using System.Windows.Markup;

namespace EliteSheets.Themes
{
    /// <summary>
    /// Loads the palettes and styles from XAML embedded in <b>this</b> assembly.
    /// Deliberately not pack://application URIs or compiled BAML (LoadComponent): both find the assembly by name,
    /// and inside Revit that can be a different EliteSheets.dll – the installed add-in, while a newer build is
    /// loaded next to it by ricaun AppLoader – which fails with "Set property Source threw an exception".
    /// </summary>
    public static class ThemeResources
    {
        public static ResourceDictionary Styles() => Load("Styles.xaml");

        public static ResourceDictionary Palette(bool dark) => Load(dark ? "DarkTheme.xaml" : "LightTheme.xaml");

        public static bool IsPalette(ResourceDictionary d) => d != null && d.Contains(PaletteMarker);

        private const string PaletteMarker = "Window.Background";

        private static ResourceDictionary Load(string file)
        {
            var assembly = typeof(ThemeResources).Assembly;
            using (var stream = assembly.GetManifestResourceStream("EliteSheets.UI.Themes." + file))
            {
                if (stream == null) throw new FileNotFoundException("Embedded theme resource not found.", file);

                string xaml;
                using (var reader = new StreamReader(stream)) xaml = reader.ReadToEnd();

                // Bind the icon namespace to the exact MaterialDesign assembly this add-in ships with (Costura),
                // not whichever version another add-in may already have loaded into Revit.
                var materialDesign = typeof(MaterialDesignThemes.Wpf.PackIcon).Assembly.FullName;
                xaml = xaml.Replace(
                    "xmlns:materialDesign=\"http://materialdesigninxaml.net/winfx/xaml/themes\"",
                    "xmlns:materialDesign=\"clr-namespace:MaterialDesignThemes.Wpf;assembly=" + materialDesign + "\"");

                return (ResourceDictionary)XamlReader.Parse(xaml);
            }
        }
    }
}
