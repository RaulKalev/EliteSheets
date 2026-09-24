using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace EliteSheets
{
    /// <summary>
    /// Minimal, explanatory motion for panel changes only (never for list rows) – shared with Sentinel:
    ///  - <c>Motion.Reveal="True"</c>: when the element becomes visible it fades in and settles from a small offset
    ///    (<c>Motion.FromY</c>, default 6 DIP). Starting from the element's current animated value
    ///    (HandoffBehavior.SnapshotAndReplace) keeps it interruptible – a new reveal never jumps.
    ///  - Reduced motion ("Animation effects" off in Windows): opacity only, no movement.
    /// WPF has no spring animation; a short critically damped-looking ease-out is used instead.
    /// </summary>
    public static class Motion
    {
        public static readonly DependencyProperty RevealProperty = DependencyProperty.RegisterAttached(
            "Reveal", typeof(bool), typeof(Motion), new PropertyMetadata(false, OnRevealChanged));

        public static readonly DependencyProperty FromYProperty = DependencyProperty.RegisterAttached(
            "FromY", typeof(double), typeof(Motion), new PropertyMetadata(6.0));

        public static bool GetReveal(DependencyObject d) => (bool)d.GetValue(RevealProperty);
        public static void SetReveal(DependencyObject d, bool value) => d.SetValue(RevealProperty, value);
        public static double GetFromY(DependencyObject d) => (double)d.GetValue(FromYProperty);
        public static void SetFromY(DependencyObject d, double value) => d.SetValue(FromYProperty, value);

        /// <summary>Windows "Animation effects" switched off.</summary>
        public static bool ReducedMotion => !SystemParameters.ClientAreaAnimation;

        private static void OnRevealChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var el = d as FrameworkElement;
            if (el == null) return;
            if ((bool)e.NewValue) el.IsVisibleChanged += OnVisibleChanged;
            else el.IsVisibleChanged -= OnVisibleChanged;
        }

        private static void OnVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var el = (FrameworkElement)sender;
            // Only animate changes after the window is up; the first layout should simply appear.
            if ((bool)e.NewValue && el.IsLoaded) Play(el);
        }

        /// <summary>Plays the reveal on any element.</summary>
        public static void Play(FrameworkElement el, double? fromY = null)
        {
            var reduced = ReducedMotion;
            var duration = TimeSpan.FromMilliseconds(reduced ? 100 : 180);
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };

            var fade = new DoubleAnimation { From = el.Opacity < 0.99 ? (double?)null : 0.0, To = 1.0, Duration = duration, EasingFunction = ease };
            el.BeginAnimation(UIElement.OpacityProperty, fade, HandoffBehavior.SnapshotAndReplace);

            if (reduced) return;
            var t = el.RenderTransform as TranslateTransform;
            if (t == null || t.IsFrozen)
            {
                t = new TranslateTransform();
                el.RenderTransform = t;
            }
            // Interrupted mid-slide: continue from the current (presentation) value instead of jumping back.
            var inFlight = Math.Abs(t.Y) > 0.01;
            var slide = new DoubleAnimation { From = inFlight ? (double?)null : (fromY ?? GetFromY(el)), To = 0, Duration = duration, EasingFunction = ease };
            t.BeginAnimation(TranslateTransform.YProperty, slide, HandoffBehavior.SnapshotAndReplace);
        }
    }
}
