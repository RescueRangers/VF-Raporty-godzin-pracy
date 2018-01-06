using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class Pracowik
    {
        public string Imie { get; set; }
        public string Nazwisko { get; set; }
        private List<Dzien> _dni = new List<Dzien>();
        public int StartIndex { get; set; }
        public int KoniecIndex { get; set; }
        public int PracownikIndex { get; set; }

        public List<Dzien> GetDni()
        {
            return _dni;
        }

        public void SetDni(List<Dzien> dni)
        {
            _dni = dni;
        }

        public void ZapelnijDni(ExcelWorksheet arkusz, List<Naglowek> naglowki)
        {
            DodajDni.DniList(PracownikIndex, StartIndex, KoniecIndex, naglowki, arkusz);
        }
    }
}