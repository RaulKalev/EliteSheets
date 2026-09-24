using EliteSheets.Services;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EliteSheets
{
    /// <summary>
    /// Shows the changelog versions the user hasn't seen yet (newest first). Built in code from the JSON so the
    /// notes keep their own wording and bullets; blank spacer notes (" ") are dropped – spacing comes from layout.
    /// </summary>
    public partial class ChangelogWindow : Window
    {
        public ChangelogWindow(Window owner, IList<UpdateLogService.VersionEntry> versions)
        {
            // Same palette as the owner (dark/light) plus the shared styles, before the XAML resolves them.
            var palette = owner?.Resources.MergedDictionaries.FirstOrDefault(Themes.ThemeResources.IsPalette)
                          ?? Themes.ThemeResources.Palette(dark: true);
            Resources.MergedDictionaries.Add(palette);
            Resources.MergedDictionaries.Add(Themes.ThemeResources.Styles());
            InitializeComponent();

            if (owner != null) Owner = owner;
            else WindowStartupLocation = WindowStartupLocation.CenterScreen;

            if (versions == null || versions.Count == 0) return;

            VersionText.Text = Loc.Format("WhatsNewVersion", versions[0].Version, versions[0].Released);

            for (int i = 0; i < versions.Count; i++)
                VersionsPanel.Children.Add(BuildVersion(versions[i], isNewest: i == 0));
        }

        private FrameworkElement BuildVersion(UpdateLogService.VersionEntry version, bool isNewest)
        {
            var panel = new StackPanel { Margin = new Thickness(0, isNewest ? 0 : 18, 0, 0) };

            // Older versions get their own heading; the newest one is already the window headline.
            if (!isNewest)
            {
                var heading = new TextBlock { Text = Loc.Format("WhatsNewOlder", version.Version) };
                heading.SetResourceReference(StyleProperty, "Type.Heading");
                panel.Children.Add(heading);
            }

            foreach (var note in version.Notes.Where(n => !string.IsNullOrWhiteSpace(n)))
            {
                var text = new TextBlock
                {
                    Text = note.Trim(),
                    Margin = new Thickness(0, 8, 0, 0),
                    TextWrapping = TextWrapping.Wrap,
                    LineHeight = 19
                };
                // Bulleted items read as content; the closing remark (no bullet) is quieter.
                text.SetResourceReference(TextBlock.ForegroundProperty,
                    note.TrimStart().StartsWith("•") ? "Text.Primary" : "Text.Secondary");
                panel.Children.Add(text);
            }

            var footer = new TextBlock
            {
                Margin = new Thickness(0, 8, 0, 0),
                Text = string.IsNullOrWhiteSpace(version.Developer)
                    ? version.Released
                    : $"{version.Released} · {version.Developer}"
            };
            footer.SetResourceReference(StyleProperty, "Type.Caption");
            panel.Children.Add(footer);

            return panel;
        }

        private void Surface_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
