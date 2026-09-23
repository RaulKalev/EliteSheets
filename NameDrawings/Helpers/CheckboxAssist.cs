using System.Windows;

namespace EliteSheets.Helpers
{
    /// <summary>
    /// Attached properties used by CustomCheckboxStyle.
    /// </summary>
    public static class CheckboxAssist
    {
        public static readonly DependencyProperty IconSizeProperty =
            DependencyProperty.RegisterAttached(
                "IconSize",
                typeof(double),
                typeof(CheckboxAssist),
                new FrameworkPropertyMetadata(30.0));

        public static double GetIconSize(DependencyObject obj) => (double)obj.GetValue(IconSizeProperty);

        public static void SetIconSize(DependencyObject obj, double value) => obj.SetValue(IconSizeProperty, value);
    }
}
