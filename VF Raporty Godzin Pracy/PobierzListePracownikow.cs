using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class PobierzListePracownikow
    {
        public static List<Pracowik> PobierzPracownikow( ExcelWorksheet arkusz)
        {
            //using (var excel = new ExcelPackage(plikExcel))
            //{
                //var arkusz = excel.Workbook.Worksheets[1];
                var pracownicy = new List<Pracowik>();
                var startWiersz = 1;
                var ostatniWiersz = arkusz.Dimension.End.Row;
                while (startWiersz<ostatniWiersz)
                {
                    var pracownik = new Pracowik();
                    for (var i = startWiersz; i < ostatniWiersz; i++)
                    {
                        if (arkusz.Cells[i,1].Value != null)
                        {
                            var nazwa = arkusz.Cells[i, 1].Value.ToString().Trim().Split(' ');
                            if (nazwa.Length == 2)
                            {
                                pracownik.Imie = nazwa[0];
                                pracownik.Nazwisko = nazwa[1];
                                pracownik.StartIndex = i;
                            }
                            else if (nazwa.Length == 3)
                            {
                                pracownik.KoniecIndex = i-1;
                                pracownicy.Add(pracownik);
                                startWiersz = i+1;
                                break;
                            }
                        }
                    }
                }

                return pracownicy;
            //}
        }
    }
}
