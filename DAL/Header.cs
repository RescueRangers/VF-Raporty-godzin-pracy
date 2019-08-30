using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DAL
{
    public class Header : INotifyPropertyChanged
    {
        private string _name;

        public string Name
        {
            get => _name;
            set
            {
                if (value == _name) return;
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public int? Column { get; set; }

        public override bool Equals(object obj)
        {
            var header = obj as Header;
            if (Name == null || header?.Name == null) return false;
            return string.Equals(Name, header.Name, System.StringComparison.OrdinalIgnoreCase);
        }

        public override string ToString()
        {
            return Name;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public Translation ToTranslate()
        {
            return new Translation
            {
                Name = Name,
                Translated = "",
                Column = Column
            };
        }
    }
}
