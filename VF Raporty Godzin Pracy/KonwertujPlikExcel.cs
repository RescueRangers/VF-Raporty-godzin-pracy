using System;
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
            else if (plikExcel.ToLower()[plikExcel.Length-1] == 'x')     
            {
                throw new FileFormatException("Niepoprawny typ pliku, oczekiwano .xls");
            }
            else
            {
                var nazwaPlikuXls = new FileInfo(plikExcel);
                var folderDocelowy = AppDomain.CurrentDomain.BaseDirectory + "Archive\\" + nazwaPlikuXls.Name;

                var sciezkPlikuDoZapisu = nazwaPlikuXls.FullName.Remove(nazwaPlikuXls.FullName.Length - 4);
                var plikXls = new Application().Workbooks;
                var arkusz = plikXls.Open(plikExcel);
                arkusz.SaveAs(sciezkPlikuDoZapisu, XlFileFormat.xlOpenXMLWorkbook);
                arkusz.Close(false);
                plikXls.Close();
                File.Delete(nazwaPlikuXls.FullName.ToString());
                return plikExcel + 'x';
            }
        }
    }
}
