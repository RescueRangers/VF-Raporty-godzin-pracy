using System.IO;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public class UtworzRaport
    {
        public Raport Stworz(string plikDoRaportu)
        {
            var plikExcel = new FileInfo(plikDoRaportu); 
            if (plikExcel.Extension.ToLower() == ".xls")
            {
                throw new FileFormatException("Niepoprawny format pliku, ozekiwano .xlsx");
            }
            else if (plikExcel == null)
            {
                throw new FileNotFoundException("Nie odnaleziono pliku.");
            }
            else
            {
                var arkuszExcel = new ExcelPackage(plikExcel).Workbook.Worksheets[1];
                var raport = new Raport(arkuszExcel);
                arkuszExcel.Dispose();
                return raport;
            }
            
        }
    }
}
