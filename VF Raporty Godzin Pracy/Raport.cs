using System.Collections.Generic;
using OfficeOpenXml;
using System.Collections.ObjectModel;
using System.Linq;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private List<Pracowik> _pracownicy = new List<Pracowik>();
        private List<Naglowek> _naglowki = new List<Naglowek>();
        private ObservableCollection<Naglowek> _niePrzetlumaczoneNaglowki = new ObservableCollection<Naglowek>();

        public List<Naglowek> TlumaczoneNaglowki { get; } = new List<Naglowek>();

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
            if (_niePrzetlumaczoneNaglowki.Count > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
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

            var tlumaczenia = serializacja.DeserializujTlumaczenia();

            var nieTlumaczoneNaglowki = _naglowki.Where(n => !tlumaczenia.Contains(n));

            var naglowkiZTlumaczeniem = _naglowki.Where(n => tlumaczenia.Contains(n));

            if (!naglowkiZTlumaczeniem.Any())
            {
                _niePrzetlumaczoneNaglowki = (ObservableCollection<Naglowek>)nieTlumaczoneNaglowki;
                return;
            }

            //var nieTlumaczoneNaglowki = new ObservableCollection<Naglowek>();
            for (int i = 0; i <= _naglowki.Count-1; i++)
            {
                var naglowek = new Naglowek();

                var naglowekDoTlumaczenia = _naglowki[i].Nazwa.ToLower();

                

                //if (tlumaczenia[i].Oryginal != _naglowki[i].Nazwa)
                //{
                //    nieTlumaczoneNaglowki.Add(_naglowki[i]);
                //    continue;
                //}
                naglowek.Nazwa = tlumaczenia[i].Przetlumaczone;
                naglowek.Kolumna = _naglowki[i].Kolumna;
                TlumaczoneNaglowki.Add(naglowek);
                //_naglowki[i].Nazwa = tlumaczenie;
            }
            _niePrzetlumaczoneNaglowki = (ObservableCollection<Naglowek>)nieTlumaczoneNaglowki;
        }

        public void CzyscListeNieprzetlumaczonych()
        {
            _niePrzetlumaczoneNaglowki.Clear();
        }

    }
}
