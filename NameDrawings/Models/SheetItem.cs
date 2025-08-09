using Autodesk.Revit.DB;
using System.ComponentModel;

namespace EliteSheets.Models
{
    public class SheetItem : INotifyPropertyChanged
    {
        private bool _isChecked;
        private string _name;
        private string _number;
        private string _viewName;
        public ElementId Id { get; set; }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string Number
        {
            get => _number;
            set { _number = value; OnPropertyChanged(nameof(Number)); }
        }

        public string ViewName
        {
            get => _viewName;
            set { _viewName = value; OnPropertyChanged(nameof(ViewName)); }
        }

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked != value)
                {
                    _isChecked = value;
                    OnPropertyChanged(nameof(IsChecked));
                }
            }
        }
        public string SheetSize { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
