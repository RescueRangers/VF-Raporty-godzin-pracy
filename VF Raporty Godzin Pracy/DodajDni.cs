using System;
using System.Collections.Generic;
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
                };
                dzien.SetGodziny(DodajGodziny.PobierzGodziny(arkusz,i,listaNaglowkow));
                dni.Add(dzien);
            }
            return dni;
        }
    }
}
