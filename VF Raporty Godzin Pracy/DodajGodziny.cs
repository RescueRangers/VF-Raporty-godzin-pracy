using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    class DodajGodziny
    {
        public static List<decimal> PobierzGodziny(ExcelWorksheet arkusz, int indeks, List<Naglowek> naglowki)
        {
            var godziny = new List<decimal>();
            foreach (var naglowek in naglowki)
            {
                if (naglowek != null) godziny.Add(Convert.ToDecimal(arkusz.Cells[indeks, (int)naglowek.Kolumna].Value));
            }
            return godziny;
        }
    }
}
