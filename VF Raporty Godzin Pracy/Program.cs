using System;
using System.IO;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    class Program
    {
        static void Main(string[] args)
        {
            var plikExcel = new FileInfo(@"d:\12.xlsx");
            var plikDoZapisu = new StreamWriter(@"d:\test.txt",false);
            var raport = new Raport();
            using (var excel = new ExcelPackage(plikExcel))
            {
                var arkusz = excel.Workbook.Worksheets[1];
                raport.ZapelnijRaport(arkusz);
            }
            //foreach (var naglowek in raport.GetNaglowki())
            //{
            //    Console.Write("{0} \t", naglowek.Nazwa);
            //}
            //foreach (var pracownik in raport.GetPracownicy())
            //{
            //    Console.WriteLine("{0} {1}", pracownik.Imie, pracownik.Nazwisko);
            //}

            foreach (var pracownik in raport.GetPracownicy())
            {
                plikDoZapisu.Write("{0} {1} \n",pracownik.Imie, pracownik.Nazwisko );
                foreach (var dzien in pracownik.GetDni())
                {
                    plikDoZapisu.Write("{0} \t",dzien.Date.Date);
                    foreach (var godzina in dzien.GetGodziny())
                    {
                        plikDoZapisu.Write(" {0:F1} \t", godzina.Warosc);
                    }
                    plikDoZapisu.Write("\n");
                }
                plikDoZapisu.Write("\n");
            }
        }
    }
}
