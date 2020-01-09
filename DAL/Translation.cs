using System;

namespace DAL
{
    [Serializable]
    public class Translation : Header
    {
        private string _translated;

        public string Translated
        {
            get => _translated;
            set
            {
                if (value == _translated) return;
                _translated = value;
                OnPropertyChanged(nameof(Translated));
            }
        }

        public Translation()
        {
            Translated = "";
        }

        public Translation(string name, string translated)
        {
            _translated = translated;
            Name = name;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}