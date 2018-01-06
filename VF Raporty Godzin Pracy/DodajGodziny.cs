using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    class DodajGodziny
    {
        public static List<Godzina> PobierzGodziny(ExcelWorksheet arkusz, int indeks, List<Naglowek> naglowki)
        {
            var godziny = new List<Godzina>();
            foreach (var naglowek in naglowki)
            {
                var godzina = new Godzina
                {
                    IndeksNaglowka = naglowek.Kolumna,
                    Warosc = Convert.ToDecimal(arkusz.Cells[indeks, naglowek.Kolumna].Value)
                };
                godziny.Add(godzina);
            }
            return godziny;
        }
    }
}
