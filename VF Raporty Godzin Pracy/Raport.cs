using System.Collections.Generic;
using OfficeOpenXml;
using System.Collections.ObjectModel;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private Dictionary<string, Pracowik> _pracownicy = new Dictionary<string, Pracowik>();
        private List<Naglowek> _naglowki = new List<Naglowek>();
        private ObservableCollection<Naglowek> _niePrzetlumaczoneNaglowki = new ObservableCollection<Naglowek>;

        public Raport(ExcelWorksheet arkusz)
        {
            _pracownicy = (PobierzListePracownikow.PobierzPracownikow(arkusz));
            _naglowki = (PobierzNaglowki.GetNaglowki(arkusz));
            foreach (var pracownik in _pracownicy)
            {
                pracownik.Value.ZapelnijDni(arkusz, _naglowki);
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

        public Dictionary<string,Pracowik> GetPracownicy()
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
                listaPracownikow.Add(pracownik.Key);
            }
            return listaPracownikow;
        }

        private void TlumaczNaglowki()
        {
            var tlumaczenia = Tlumacz.LadujTlumaczenia();
            var nieTlumaczoneNaglowki = new ObservableCollection<Naglowek>();
            for (int i = 0; i <= _naglowki.Count-1; i++)
            {
                var naglowekDoTlumaczenia = _naglowki[i].Nazwa.ToLower();

                if (tlumaczenia.TryGetValue(naglowekDoTlumaczenia, out string tlumaczenie) == false)
                {
                    nieTlumaczoneNaglowki.Add(_naglowki[i]);
                    continue;
                }
                _naglowki[i].Nazwa = tlumaczenie;
            }
            _niePrzetlumaczoneNaglowki = nieTlumaczoneNaglowki;
        }
    }
}
