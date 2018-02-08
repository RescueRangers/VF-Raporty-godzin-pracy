using System.Collections.Generic;
using OfficeOpenXml;
using System.Collections.ObjectModel;
using System.Linq;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private List<Pracowik> _pracownicy;
        private List<Naglowek> _naglowki;

        private ObservableCollection<Naglowek> _niePrzetlumaczoneNaglowki = new ObservableCollection<Naglowek>();

        public List<Pracowik> Pracownicy
        {
            get => _pracownicy;
            set => _pracownicy = value;
        }


        public ObservableCollection<Naglowek> NiePrzetlumaczoneNaglowki
        {
            get => _niePrzetlumaczoneNaglowki;
            set => _niePrzetlumaczoneNaglowki = value;
        }

        public List<Naglowek> TlumaczoneNaglowki = new List<Naglowek>();

        public Raport(ExcelWorksheet arkusz)
        {
            _pracownicy = (PobierzListePracownikow.PobierzPracownikow(arkusz));
            _naglowki = (PobierzNaglowki.GetNaglowki(arkusz));
            foreach (var pracownik in _pracownicy)
            {
                pracownik.ZapelnijDni(arkusz, _naglowki);
            }
            TlumaczNaglowki();
        }

        public bool CzyPrzetlumaczoneNaglowki()
        {
            return !_niePrzetlumaczoneNaglowki.Any();
        }

        public ObservableCollection<Naglowek> PobierzNiePrzetlumaczoneNaglowki()
        {
            return _niePrzetlumaczoneNaglowki;
        }

        public List<Pracowik> GetPracownicy()
        {
            return _pracownicy;
        }

        public List<Naglowek> GetNaglowki()
        {
            return _naglowki;
        }

        public List<string> GetNazwyPracownikow()
        {
            var listaPracownikow = new List<string>();
            foreach (var pracownik in _pracownicy)
            {
                listaPracownikow.Add($"{pracownik.Nazwisko} {pracownik.Imie}");
            }
            return listaPracownikow;
        }

        public ObservableCollection<Pracowik> PobierzPracownikowDoWidoku()
        {
            var pracownicy = new ObservableCollection<Pracowik>();
            foreach (var pracownik in _pracownicy)
            {
                pracownicy.Add(pracownik);
            }

            return pracownicy;
        }

        public void TlumaczNaglowki()
        {
            var serializacja = new SerializacjaTlumaczen();

            TlumaczoneNaglowki.Clear();
            TlumaczoneNaglowki = _naglowki;

            var tlumaczenia = serializacja.DeserializujTlumaczenia();

            var nieTlumaczoneNaglowki = new ObservableCollection<Naglowek>(_naglowki.Where(n => !tlumaczenia.Contains(n)).ToList());
            var tlumaczoneNaglowki = tlumaczenia.Where(t => TlumaczoneNaglowki.Contains(t)).ToList();

            if (!tlumaczoneNaglowki.Any())
            {
                _niePrzetlumaczoneNaglowki = nieTlumaczoneNaglowki;
                return;
            }

            foreach (var naglowek in tlumaczoneNaglowki)
            {
                var indeksNaglowka = TlumaczoneNaglowki.FindIndex(n => n.Equals(naglowek));
                TlumaczoneNaglowki[indeksNaglowka].Nazwa = naglowek.Przetlumaczone;
            }

            
            _niePrzetlumaczoneNaglowki = nieTlumaczoneNaglowki;
        }

        public void CzyscListeNieprzetlumaczonych()
        {
            _niePrzetlumaczoneNaglowki.Clear();
        }

        public void DodajTlumaczenie(Tlumaczenie tlumaczenie)
        {
            var naglowek = new Naglowek
            {
                Nazwa = tlumaczenie.Przetlumaczone,
                Kolumna = tlumaczenie.Kolumna
            };
            TlumaczoneNaglowki.Remove(tlumaczenie);
            TlumaczoneNaglowki.Add(naglowek);
            _niePrzetlumaczoneNaglowki.Remove(tlumaczenie);
        }
    }
}
