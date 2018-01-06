﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    class Raport
    {
        private List<Pracowik> _pracownicy = new List<Pracowik>();
        private List<Naglowek> _naglowki = new List<Naglowek>();

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

        public void ZapelnijRaport(ExcelWorksheet arkusz)
        {
            SetPracownicy(PobierzListePracownikow.PobierzPracownikow(arkusz));
            SetNaglowki(PobierzNaglowki.GetNaglowki(arkusz));
            foreach (var pracownik in GetPracownicy())
            {
                pracownik.ZapelnijDni(arkusz,GetNaglowki());
            }
        }
    }
}
