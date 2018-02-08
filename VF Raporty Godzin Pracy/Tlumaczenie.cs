using System;
using System.ComponentModel;

namespace VF_Raporty_Godzin_Pracy
{
    [Serializable]
    public class Tlumaczenie : Naglowek
    {
        private string _przetlumaczone;

        public string Przetlumaczone
        {
            get => _przetlumaczone;
            set
            {
                if (value == _przetlumaczone) return;
                _przetlumaczone = value;
                OnPropertyChanged(nameof(Przetlumaczone));
            }
        }

        public Tlumaczenie()
        {
            Kolumna = null;
        }

        public override string ToString()
        {
            return Nazwa;
        }
    }
}
