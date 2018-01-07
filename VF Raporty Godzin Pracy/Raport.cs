using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private List<Pracowik> _pracownicy;
        private List<Naglowek> _naglowki;

        public Raport(ExcelWorksheet arkusz)        
        {
            _pracownicy=(PobierzListePracownikow.PobierzPracownikow(arkusz));
            _naglowki=(PobierzNaglowki.GetNaglowki(arkusz));
            foreach (var pracownik in _pracownicy)
            {
                pracownik.ZapelnijDni(arkusz,_naglowki);
            }
        }

        public List<Pracowik> GetPracownicy()
        {
            return _pracownicy;
        }

        public void SetPracownicy(List<Pracowik> pracownicy)
        {
            _pracownicy = pracownicy;
        }

        public List<Naglowek> GetNaglowki()
        {
            return _naglowki;
        }

        public void SetNaglowki(List<Naglowek> naglowki)
        {
            _naglowki = naglowki;
        }


    }
}
