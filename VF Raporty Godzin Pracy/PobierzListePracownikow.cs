using System.Collections.Generic;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class PobierzListePracownikow
    {
        public static List<Pracowik> PobierzPracownikow( ExcelWorksheet arkusz)
        {
            var pracownicy = new List<Pracowik>();
            var startWiersz = 1;
            var ostatniWiersz = arkusz.Dimension.End.Row;
            var j = 0;
            while (startWiersz < ostatniWiersz)
            {
                var pracownik = new Pracowik();
                for (var i = startWiersz; i < ostatniWiersz; i++)
                {
                    if (arkusz.Cells[i, 1].Value != null)
                    {
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
            }
            return pracownicy;
        }
    }
}
