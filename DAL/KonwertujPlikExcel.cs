using System;
using System.IO;
using System.Net.Mime;
using Microsoft.Office.Interop.Excel;

namespace DAL
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
                throw new Exception("Niepoprawny typ pliku, oczekiwano .xls");
            }

            var nazwaPlikuXls = new FileInfo(plikExcel);

            var sciezkPlikuDoZapisu = nazwaPlikuXls.FullName.Remove(nazwaPlikuXls.FullName.Length - 4);
            var plikXls = new Application().Workbooks;
            var arkusz = plikXls.Open(plikExcel);
            arkusz.SaveAs(sciezkPlikuDoZapisu + ".xlsx", XlFileFormat.xlOpenXMLWorkbook);
            arkusz.Close(false);
            plikXls.Close();
            File.Delete(nazwaPlikuXls.FullName);
            return plikExcel + 'x';
        }
    }
}
