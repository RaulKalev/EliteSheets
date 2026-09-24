using EliteSheets.Services;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace EliteSheets.Helpers
{
    /// <summary>
    /// XAML lookup for translated strings: Content="{helpers:Tr Print}".
    /// Bound to <see cref="LocSource"/>, so text updates live when the language changes.
    /// </summary>
    [MarkupExtensionReturnType(typeof(object))]
    public class TrExtension : MarkupExtension
    {
        public TrExtension() { }

        public TrExtension(string key)
        {
            Key = key;
        }

        [ConstructorArgument("key")]
        public string Key { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            // Plain CLR properties (e.g. Binding.StringFormat) can't hold a binding - give them the current text
            var target = serviceProvider?.GetService(typeof(IProvideValueTarget)) as IProvideValueTarget;
            if (!(target?.TargetProperty is DependencyProperty))
                return Loc.Get(Key);

            var binding = new Binding($"[{Key}]") { Source = LocSource.Instance, Mode = BindingMode.OneWay };
            return binding.ProvideValue(serviceProvider);
        }
    }

    /// <summary>
    /// Bindable view of <see cref="Loc"/>; raises a change for every key when the language switches.
    /// </summary>
    public sealed class LocSource : INotifyPropertyChanged
    {
        public static LocSource Instance { get; } = new LocSource();

        private LocSource()
        {
            Loc.LanguageChanged += (s, e) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Binding.IndexerName));
        }

        public string this[string key] => Loc.Get(key);

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
