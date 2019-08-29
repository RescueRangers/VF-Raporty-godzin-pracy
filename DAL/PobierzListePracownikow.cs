using System.Collections.Generic;
using OfficeOpenXml;

namespace DAL
{
    public static class PobierzListePracownikow
    {
        public static List<Pracownik> PobierzPracownikow( ExcelWorksheet arkusz)
        {
            var pracownicy = new List<Pracownik>();
            var startWiersz = 1;
            var ostatniWiersz = arkusz.Dimension.End.Row;
            var j = 0;
            while (startWiersz < ostatniWiersz)
            {
                var pracownik = new Pracownik();
                for (var i = startWiersz; i < ostatniWiersz; i++)
                {
                    if (arkusz.Cells[i, 1].Value == null) continue;
                    var nazwa = arkusz.Cells[i, 1].Value.ToString().Trim().Split(' ');
                    if (nazwa[nazwa.Length-1].ToLower() != "total")
                    {
                        pracownik.Imie = nazwa[0];
                        pracownik.Nazwisko = nazwa[1];
                        pracownik.UstawStartIndeks(i);
                    }
                    else
                    {
                        pracownik.UstawKoniecIndeks(i);
                        pracownik.UstawPracownikIndeks(j);
                        pracownicy.Add(pracownik);
                        j++;
                        startWiersz = i + 1;
                        break;
                    }
                }
            }
            return pracownicy;
        }
    }
}
