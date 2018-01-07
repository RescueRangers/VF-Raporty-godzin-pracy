using System.IO;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    class Program
    {
        static void Main()
        {
            var plikExcel = new FileInfo(@"d:\12.xlsx");
            var plikDoZapisu = new StreamWriter(@"d:\test.txt",false);
            var excel = new ExcelPackage(plikExcel);
            var arkusz = excel.Workbook.Worksheets[1];
            var raport = new Raport(arkusz);
            excel.Dispose();
            ZapiszExcel.ZapiszDoExcel(raport);
            foreach (var pracownik in raport.GetPracownicy())
            {
                var nazwaPliku = string.Format(@"d:\test\{0} {1}.txt",pracownik.Nazwisko,pracownik.Imie);
                var zapisDoPliku = new StreamWriter(nazwaPliku, true);
                zapisDoPliku.Write("{0} {1} \n",pracownik.Imie, pracownik.Nazwisko );
                foreach (var dzien in pracownik.GetDni())
                {
                    zapisDoPliku.Write("{0} \t",dzien.Date.Date);
                    foreach (var godzina in dzien.GetGodziny())
                    {
                        zapisDoPliku.Write(" {0:F1} \t", godzina);
                    }
                    zapisDoPliku.Write("\n");
                }
                zapisDoPliku.Write("\n");
            }

            foreach (var pracownik in raport.GetPracownicy())
            {
                plikDoZapisu.Write("{0} {1} \n",pracownik.Imie, pracownik.Nazwisko );
                foreach (var dzien in pracownik.GetDni())
                {
                    plikDoZapisu.Write("{0} \t",dzien.Date.Date);
                    foreach (var godzina in dzien.GetGodziny())
                    {
                        plikDoZapisu.Write(" {0:F1} \t", godzina);
                    }
                    plikDoZapisu.Write("\n");
                }
                plikDoZapisu.Write("\n");
            }
        }
    }
}
