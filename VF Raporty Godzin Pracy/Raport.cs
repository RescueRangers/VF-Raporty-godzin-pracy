using System.Collections.Generic;
using OfficeOpenXml;
using System.Linq;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private List<Pracowik> _pracownicy;
        private List<Naglowek> _naglowki;


        private List<Naglowek> _niePrzetlumaczoneNaglowki = new List<Naglowek>();

        public List<Pracowik> Pracownicy
        {
            get => _pracownicy;
            set => _pracownicy = value;
        }


        public List<Naglowek> NiePrzetlumaczoneNaglowki
        {
            get => _niePrzetlumaczoneNaglowki;
            set => _niePrzetlumaczoneNaglowki = value;
        }

        public List<Naglowek> TlumaczoneNaglowki { get; set; }

        public Raport(ExcelWorksheet arkusz)
        {
            TlumaczoneNaglowki = new List<Naglowek>();
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

        public List<string> GetNazwyPracownikow()
        {
            var listaPracownikow = new List<string>();
            foreach (var pracownik in _pracownicy)
            {
                listaPracownikow.Add($"{pracownik.Nazwisko} {pracownik.Imie}");
            }
            return listaPracownikow;
        }

        public void TlumaczNaglowki()
        {
            var serializacja = new SerializacjaTlumaczen();

            TlumaczoneNaglowki.Clear();
            TlumaczoneNaglowki = _naglowki;

            var tlumaczenia = serializacja.DeserializujTlumaczenia();

            var nieTlumaczoneNaglowki = new List<Naglowek>(_naglowki.Where(n => !tlumaczenia.Contains(n)));
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
