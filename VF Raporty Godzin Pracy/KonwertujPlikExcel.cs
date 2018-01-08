using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace VF_Raporty_Godzin_Pracy
{
    public class KonwertujPlikExcel
    {
        public static string XlsDoXlsx(string plikExcel)
        {
            if (plikExcel == null)
            {
                throw new FileNotFoundException("Nie znaleziono pliku");
            }
            else if (plikExcel.ToLower()[plikExcel.Count()] == 'x')     
            {
                throw new FileFormatException("Niepoprawny typ pliku, oczekiwano .xls");
            }
            else
            {
                var nazwaPlikuXls = new FileInfo(plikExcel);
                
                var plikXls = new Application().Workbooks.Open(plikExcel);
                plikXls.SaveAs(nazwaPlikuXls.Name.Remove(nazwaPlikuXls.Name.Count() - 4), "xlOpenXMLWorkbook");
                plikXls.Close(false);
                return plikExcel + 's';
            }
        }
    }
}
