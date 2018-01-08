using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private Dictionary<string, Pracowik> _pracownicy = new Dictionary<string, Pracowik>();
        private List<Naglowek> _naglowki = new List<Naglowek>();

        public Raport(ExcelWorksheet arkusz)
        {
            _pracownicy = (PobierzListePracownikow.PobierzPracownikow(arkusz));
            _naglowki = (PobierzNaglowki.GetNaglowki(arkusz));
            foreach (var pracownik in _pracownicy)
            {
                pracownik.Value.ZapelnijDni(arkusz, _naglowki);
            }
        }

        public Dictionary<string,Pracowik> GetPracownicy()
        {
            return _pracownicy;
        }

        public List<Naglowek> GetNaglowki()
        {
            return _naglowki;
        }

        public void SetNaglowki(List<Naglowek> naglowki)
        {
            _naglowki = naglowki;
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
    }
}
