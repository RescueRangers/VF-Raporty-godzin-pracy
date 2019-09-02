using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace DAL
{
    public class Report
    {
        public List<Employee> Employees { get; set; }

        public List<Day> NotTranslatedHeaders { get; set; } = new List<Day>();
        public List<Header> TranslatedHeaders { get; set; }
        public List<Header> Headers { get; private set; }

        public Report(ExcelWorksheet worksheet)
        {
            TranslatedHeaders = new List<Header>();
            Employees = GetEmployees(worksheet);
            Headers = GetHeaders(worksheet);
            foreach (var employee in Employees)
            {
                employee.FillDays(worksheet, Headers);
            }
            TranslateHeaders();
        }

        private static List<Employee> GetEmployees( ExcelWorksheet arkusz)
        {
            var employees = new List<Employee>();
            var firstRow = 1;
            var lastRow = arkusz.Dimension.End.Row;
            var j = 0;
            while (firstRow < lastRow)
            {
                var employee = new Employee();
                for (var i = firstRow; i < lastRow; i++)
                {
                    if (arkusz.Cells[i, 1].Value == null) continue;
                    var name = arkusz.Cells[i, 1].Value.ToString().Trim().Split(' ');
                    if (!string.Equals(name[name.Length - 1], "total", System.StringComparison.OrdinalIgnoreCase))
                    {
                        employee.FirstName = name[0];
                        employee.Lastname = name[1];
                        employee.SetStartIndex(i);
                    }
                    else
                    {
                        employee.SetEndIndex(i);
                        employees.Add(employee);
                        j++;
                        firstRow = i + 1;
                        break;
                    }
                }
            }
            return employees;
        }

        private static List<Header> GetHeaders(ExcelWorksheet arkusz)
        {
            var headers = new List<Header>();
            var lastColumn = arkusz.Dimension.End.Column;
            for (var i = 1; i < lastColumn; i++)
            {
                if (arkusz.Cells[6, i].Value == null || string.Equals(arkusz.Cells[6, i].Value.ToString(),
                        "grand total", System.StringComparison.OrdinalIgnoreCase)) continue;
                var header = new Header {Column = i, Name = arkusz.Cells[6, i].Value.ToString()};
                headers.Add(header);
            }
            return headers;
        }

        public static Report Create(string reportFile)
        {
            var excelFile = new FileInfo(reportFile); 
            if (string.Equals(excelFile.Extension, ".xls", System.StringComparison.OrdinalIgnoreCase))
            {
                throw new FileLoadException("Incorrect file extension");
            }

            using (var excelWorksheet = new ExcelPackage(excelFile).Workbook.Worksheets[1])
            {
                if (excelWorksheet.Cells[1, 2].Text != "Department Code") throw new FileLoadException("Incorrect report file");
                return new Report(excelWorksheet);
            }
        }

        public bool AreHeadersTranslated()
        {
            return !NotTranslatedHeaders.Any();
        }

        public void TranslateHeaders()
        {
            var serialization = new TranslationSerialization();

            TranslatedHeaders.Clear();
            TranslatedHeaders = Headers.Select(header => new Header()
            {
                Column = header.Column,
                Name = header.Name
            }).ToList();

            var translations = serialization.DeserializeTranslations();
            var untranslatedAbsences = new List<Day>();

            foreach (var absence in Employees.SelectMany(d => d.Days).Where(d => d.WorkType == WorkType.Absence && string.IsNullOrWhiteSpace(d.TranslatedAbsence)))
            {
                var index = translations.FindIndex(h => h.Name == absence.Absence);
                if (index == -1)
                {
                    untranslatedAbsences.Add(absence);
                    continue;
                }
                absence.TranslatedAbsence = translations[index].Translated;
            }

            NotTranslatedHeaders = untranslatedAbsences;
        }
    }
}
