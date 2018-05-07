using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Globalization;

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
                    Date = DateTime.ParseExact(arkusz.Cells[i, 7].Text,"dd-MM-yyyy",new CultureInfo("pl-PL"))
                };
                dzien.SetGodziny(DodajGodziny.PobierzGodziny(arkusz,i,listaNaglowkow));
                dni.Add(dzien);
            }
            return dni;
        }
    }
}
