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
            InitializeComponent();

            Owner = owner;
            Title = title;
            TitleText.Text = title;
            MessageText.Text = message;

            // Pick up the owner's currently loaded theme dictionary (DarkTheme/LightTheme)
            var theme = owner?.Resources.MergedDictionaries.LastOrDefault();
            if (theme != null)
                Resources.MergedDictionaries.Add(theme);

            switch (kind)
            {
                case ThemedDialogKind.Success:
                    DialogIcon.Kind = PackIconKind.CheckCircleOutline;
                    DialogIcon.Foreground = new SolidColorBrush(Color.FromRgb(0x20, 0xC9, 0x97));
                    break;
                case ThemedDialogKind.Error:
                    DialogIcon.Kind = PackIconKind.CloseCircleOutline;
                    DialogIcon.Foreground = new SolidColorBrush(Color.FromRgb(0xDC, 0x6B, 0x6B));
                    break;
                case ThemedDialogKind.Info:
                    DialogIcon.Kind = PackIconKind.InformationOutline;
                    DialogIcon.Foreground = new SolidColorBrush(Color.FromRgb(0x70, 0xBA, 0xBC));
                    break;
            }

            if (yesNo)
            {
                YesButton.Visibility = Visibility.Visible;
                YesButton.IsDefault = true;
                OkButton.IsDefault = false;
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

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
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
