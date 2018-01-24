using System.IO;
using OfficeOpenXml;

namespace VF_Raporty_Godzin_Pracy
{
    public static class UtworzRaport
    {
        public static Raport Stworz(string plikDoRaportu)
        {
            var plikExcel = new FileInfo(plikDoRaportu); 
            if (plikExcel.Extension.ToLower() == ".xls")
            {
                return null;
                //throw new FileFormatException("Niepoprawny format pliku, ozekiwano .xlsx");
            }
            else if (plikExcel == null)
            {
                return null;
                //throw new FileNotFoundException("Nie odnaleziono pliku.");
            }
            else
            {
                var arkuszExcel = new ExcelPackage(plikExcel).Workbook.Worksheets[1];
                if (arkuszExcel.Cells[1,2].Text == "Department Code")
                {
                    var raport = new Raport(arkuszExcel);
                    arkuszExcel.Dispose();
                    return raport;
                }
                return null;
            }
            
        }
    }
}
