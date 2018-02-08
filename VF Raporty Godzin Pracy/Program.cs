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
            var arkuszExcel = new ExcelPackage(plikExcel).Workbook.Worksheets[1];
            var raport = new Raport(arkuszExcel);
            arkuszExcel.Dispose();
            ZapiszExcel.ZapiszDoExcel(raport);
            
        }
    }
}
