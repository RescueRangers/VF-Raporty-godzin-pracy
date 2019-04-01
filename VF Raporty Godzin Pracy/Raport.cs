﻿using System.Collections.Generic;
using OfficeOpenXml;
using System.Linq;

namespace VF_Raporty_Godzin_Pracy
{
    public class Raport
    {
        private List<Pracownik> _pracownicy;
        private List<Naglowek> _naglowki;


        private List<Naglowek> _niePrzetlumaczoneNaglowki = new List<Naglowek>();

        public List<Pracownik> Pracownicy
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
        public List<Naglowek> Naglowki { get => _naglowki;
            private set => _naglowki = value; }

        public Raport(ExcelWorksheet arkusz)
        {
            TlumaczoneNaglowki = new List<Naglowek>();
            _pracownicy = (PobierzListePracownikow.PobierzPracownikow(arkusz));
            Naglowki = (PobierzNaglowki.GetNaglowki(arkusz));
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

        public void TlumaczNaglowki()
        {
            var serializacja = new SerializacjaTlumaczen();

            TlumaczoneNaglowki.Clear();
            TlumaczoneNaglowki = Naglowki.Select(naglowek => new Naglowek()
            {
                Kolumna = naglowek.Kolumna,
                Nazwa = naglowek.Nazwa
            }).ToList();

            var tlumaczenia = serializacja.DeserializujTlumaczenia();

            var nieTlumaczoneNaglowki = new List<Naglowek>(Naglowki.Where(n => !tlumaczenia.Contains(n)));
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
    }
}
