using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class Pracowik
    {
        public string Nazwisko { get; set; }
        public string Imie { get; set; }
        private List<Dzien> _dni = new List<Dzien>();
        private int _startIndex;
        private int _koniecIndex;
        private int _pracownikIndex;

        public List<Dzien> GetDni()
        {
            return _dni;
        }

        public void ZapelnijDni(ExcelWorksheet arkusz, List<Naglowek> naglowki)
        {
            _dni = DodajDni.DniList(_pracownikIndex, _startIndex, _koniecIndex, naglowki, arkusz);
        }

        /// <summary>
        /// Zwraca string z nazwiskiem i imieniem pracownika
        /// </summary>
        /// <returns></returns>
        public string NazwaPracownika()
        {
            return $"{Nazwisko} {Imie}";
        }

        /// <summary>
        /// Zwracy daty
        /// </summary>
        /// <returns></returns>
        public List<DateTime> GetDaty()
        {
            if (_dni.Count == 0) throw new InvalidOperationException("Lista dni jest pusta");
            var daty = new List<DateTime>();
            foreach (var dzien in _dni)
            {
                daty.Add(dzien.Date);
            }
            return daty;

        }

        public void UstawStartIndeks(int startIndeks)
        {
            _startIndex = startIndeks;
        }

        public void UstawKoniecIndeks(int koniecIndeks)
        {
            _koniecIndex = koniecIndeks;
        }

        public void UstawPracownikIndeks(int pracownikIndeks)
        {
            _pracownikIndex = pracownikIndeks;
        }
    }
}