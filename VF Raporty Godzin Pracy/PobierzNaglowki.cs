using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    static class PobierzNaglowki
    {
        public static List<Naglowek> GetNaglowki(ExcelWorksheet arkusz)
        {
            var naglowki = new List<Naglowek>();
            var ostatniaKolumna = arkusz.Dimension.End.Column;
            for (var i = 1; i < ostatniaKolumna; i++)
            {
                if (arkusz.Cells[6, i].Value != null)
                {
                    var naglowek = new Naglowek {Kolumna = i, Nazwa = arkusz.Cells[6, i].Value.ToString()};
                    naglowki.Add(naglowek);
                }
            }
            return naglowki;
        }
    }
}
