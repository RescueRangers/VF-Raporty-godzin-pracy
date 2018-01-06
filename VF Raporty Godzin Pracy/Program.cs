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
            var plikDoZapisu = new StreamWriter(@"d:\test.txt",true);
            var raport = new Raport();
            using (var excel = new ExcelPackage(plikExcel))
            {
                var arkusz = excel.Workbook.Worksheets[1];
                raport.ZapelnijRaport(arkusz);
            }
            foreach (var naglowek in raport.GetNaglowki())
            {
                Console.Write("{0} \t", naglowek.Nazwa);
            }
            foreach (var pracownik in raport.GetPracownicy())
            {
                Console.WriteLine("{0} {1}", pracownik.Imie, pracownik.Nazwisko);
            }
        }
    }
}
