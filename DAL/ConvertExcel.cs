using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace DAL
{
    public static class ConvertExcel
    {
        public static string XlsToXlsx(string excelFile)
        {
            if (excelFile == null)
            {
                throw new FileNotFoundException("Nie znaleziono pliku");
            }
            if (excelFile.ToLower()[excelFile.Length-1] == 'x')
            {
                throw new Exception("Niepoprawny typ pliku, oczekiwano .xls");
            }

            var xlsFileName = new FileInfo(excelFile);

            var savePath = xlsFileName.FullName.Remove(xlsFileName.FullName.Length - 4);
            var xlsFile = new Application().Workbooks;
            var workbook = xlsFile.Open(excelFile);
            workbook.SaveAs(savePath + ".xlsx", XlFileFormat.xlOpenXMLWorkbook);
            workbook.Close(false);
            xlsFile.Close();
            File.Delete(xlsFileName.FullName);
            return excelFile + 'x';
        }
    }
}