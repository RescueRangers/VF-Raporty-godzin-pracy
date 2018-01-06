using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class DodajDni
    {
        public static List<Dzien> DniList(int indeksPracownika, int pracownikStart, int pracownikKoniec, List<Naglowek> listaNaglowkow, ExcelWorksheet arkusz)
        {
            var dni = new List<Dzien>();
            for (var i = pracownikStart; i < pracownikKoniec; i++)
            {
                var dzien = new Dzien
                {
                    Date = DateTime.Parse(arkusz.Cells[i, 7].Text),
                    Godziny = DodajGodziny.PobierzGodziny(arkusz, i, listaNaglowkow)
                };
                dni.Add(dzien);
            }
            Console.WriteLine("Skonczono {0} z {1}",indeksPracownika, 45);
            return dni;
        }
    }
}
