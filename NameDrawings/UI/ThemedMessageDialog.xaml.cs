using EliteSheets.Services;
using MaterialDesignThemes.Wpf;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Input;
using System.Windows.Media;

namespace EliteSheets
{
    public enum ThemedDialogKind
    {
        Info,
        Success,
        Warning,
        Error
    }

    /// <summary>
    /// OK or Yes/No dialog that reuses the owner window's light/dark theme brushes
    /// and always opens on top of the owner.
    /// </summary>
    public partial class ThemedMessageDialog : Window
    {
        private ThemedMessageDialog(Window owner, string title, string message, ThemedDialogKind kind, bool yesNo)
        {
            // Same palette as the owner (dark/light) plus the shared styles, merged before the XAML resolves them.
            var palette = owner?.Resources.MergedDictionaries.FirstOrDefault(Themes.ThemeResources.IsPalette)
                          ?? Themes.ThemeResources.Palette(dark: true);
            Resources.MergedDictionaries.Add(palette);
            Resources.MergedDictionaries.Add(Themes.ThemeResources.Styles());
            InitializeComponent();

            Owner = owner;
            Title = title;
            TitleText.Text = title;
            MessageText.Text = message;

            // Glyph shape + status colour, so the kind never relies on colour alone
            switch (kind)
            {
                case ThemedDialogKind.Success:
                    DialogIcon.Kind = PackIconKind.CheckCircleOutline;
                    DialogIcon.SetResourceReference(ForegroundProperty, "Status.Ok");
                    break;
                case ThemedDialogKind.Error:
                    DialogIcon.Kind = PackIconKind.CloseCircleOutline;
                    DialogIcon.SetResourceReference(ForegroundProperty, "Status.Error");
                    break;
                case ThemedDialogKind.Info:
                    DialogIcon.Kind = PackIconKind.InformationOutline;
                    DialogIcon.SetResourceReference(ForegroundProperty, "Status.Info");
                    break;
                default:
                    DialogIcon.Kind = PackIconKind.AlertOutline;
                    DialogIcon.SetResourceReference(ForegroundProperty, "Status.Warning");
                    break;
            }

            if (yesNo)
            {
                // "No" becomes the secondary button on the left; "Yes" is the primary, default action.
                YesButton.Visibility = Visibility.Visible;
                YesButton.IsDefault = true;
                OkButton.IsDefault = false;
                OkButton.SetResourceReference(StyleProperty, "Button.Secondary");
                OkButton.Content = Loc.Get("No");
                AutomationProperties.SetName(OkButton, Loc.Get("No"));
            }
        }

        public static void Show(Window owner, string title, string message,
            ThemedDialogKind kind = ThemedDialogKind.Warning)
        {
            ShowOnTop(new ThemedMessageDialog(owner, title, message, kind, false));
        }

        /// <summary>Returns true when the user clicks Yes.</summary>
        public static bool AskYesNo(Window owner, string title, string message,
            ThemedDialogKind kind = ThemedDialogKind.Info)
        {
            return ShowOnTop(new ThemedMessageDialog(owner, title, message, kind, true)) == true;
        }

        private static bool? ShowOnTop(ThemedMessageDialog dialog)
        {
            var owner = dialog.Owner;
            if (owner != null)
            {
                if (owner.WindowState == WindowState.Minimized)
                    owner.WindowState = WindowState.Normal;
                owner.Activate();
            }
            else
            {
                dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                dialog.Topmost = true;
            }

            dialog.Loaded += (s, e) => dialog.Activate();
            return dialog.ShowDialog();
        }

        /// <summary>The whole surface is the drag handle (there is no separate title bar).</summary>
        private void Surface_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void YesButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
