using System.IO;
using Microsoft.Office.Interop.Excel;

namespace VF_Raporty_Godzin_Pracy
{
    public static class KonwertujPlikExcel
    {
        public static string XlsDoXlsx(string plikExcel)
        {
            if (plikExcel == null)
            {
                throw new FileNotFoundException("Nie znaleziono pliku");
            }
            if (plikExcel.ToLower()[plikExcel.Length-1] == 'x')     
            {
                throw new FileFormatException("Niepoprawny typ pliku, oczekiwano .xls");
            }
            
            var nazwaPlikuXls = new FileInfo(plikExcel);

            var sciezkPlikuDoZapisu = nazwaPlikuXls.FullName.Remove(nazwaPlikuXls.FullName.Length - 4);
            var plikXls = new Application().Workbooks;
            var arkusz = plikXls.Open(plikExcel);
            arkusz.SaveAs(sciezkPlikuDoZapisu, XlFileFormat.xlOpenXMLWorkbook);
            arkusz.Close(false);
            plikXls.Close();
            File.Delete(nazwaPlikuXls.FullName);
            return plikExcel + 'x';
            
        }
    }
}
