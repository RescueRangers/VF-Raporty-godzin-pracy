using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class Pracowik
    {
        public string Imie;
        public string Nazwisko;
        public List<Dzien> Dni;
        public int StartIndex;
        public int KoniecIndex;
        public int PracownikIndex;

        public Pracowik()
        {
            Dni = new List<Dzien>();
        }

        public void ZapelnijDni(ExcelWorksheet arkusz, List<Naglowek> naglowki)
        {
            DodajDni.DniList(PracownikIndex, StartIndex, KoniecIndex, naglowki, arkusz);
        }
    }
}