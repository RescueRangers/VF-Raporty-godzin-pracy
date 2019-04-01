using System.ComponentModel;
using System.Runtime.CompilerServices;
using VF_Raporty_Godzin_Pracy.Annotations;

namespace VF_Raporty_Godzin_Pracy
{
    public class Naglowek : INotifyPropertyChanged
    {

        private string _nazwa;

        public string Nazwa
        {
            get => _nazwa;
            set
            {
                if (value == _nazwa) return;
                _nazwa = value;
                OnPropertyChanged(nameof(Nazwa));
            }
        }
        public int? Kolumna { get; set; }

        public override bool Equals(object obj)
        {
            var naglowek = obj as Naglowek;
            if (Nazwa == null || naglowek?.Nazwa == null) return false;
            return Nazwa.ToLower() == naglowek.Nazwa.ToLower();
        }

        public override string ToString()
        {
            return Nazwa;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public Tlumaczenie DoTlumaczenia()
        {
            return new Tlumaczenie
            {
                Nazwa = Nazwa,
                Przetlumaczone = "",
                Kolumna = Kolumna
            };
        }
    }
}
