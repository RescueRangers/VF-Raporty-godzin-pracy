using System.Collections.Generic;
using System.IO;
using System.Linq;
using DAL.Interfaces;
using OfficeOpenXml;

namespace DAL
{
    public class Report : IReport
    {
        public List<Employee> Employees { get; set; }

        public List<Day> NotTranslatedHeaders { get; private set; } = new List<Day>();
        private List<Header> _translatedHeaders;
        public List<Header> Headers { get; }

        public Report(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return;
            }

            _translatedHeaders = new List<Header>();
            Employees = GetEmployees(worksheet);
            Headers = GetHeaders(worksheet);
            foreach (var employee in Employees)
            {
                employee.FillDays(worksheet, Headers);
            }
            TranslateHeaders();
        }

        private static List<Employee> GetEmployees(ExcelWorksheet worksheet)
        {
            var employees = new List<Employee>();
            var firstRow = 1;
            var lastRow = worksheet.Dimension.End.Row;
            while (firstRow < lastRow)
            {
                var employee = new Employee();
                for (var i = firstRow; i < lastRow; i++)
                {
                    if (worksheet.Cells[i, 1].Value == null) continue;
                    var name = worksheet.Cells[i, 1].Value.ToString().Trim().Split(' ');
                    if (!string.Equals(name[name.Length - 1], "total", System.StringComparison.OrdinalIgnoreCase))
                    {
                        employee.FirstName = name[0];
                        employee.LastName = name[1];
                        employee.SetStartIndex(i);
                    }
                    else
                    {
                        employee.SetEndIndex(i);
                        employees.Add(employee);
                        firstRow = i + 1;
                        break;
                    }
                }
            }
            return employees;
        }

        private static List<Header> GetHeaders(ExcelWorksheet worksheet)
        {
            var headers = new List<Header>();
            var lastColumn = worksheet.Dimension.End.Column;
            for (var i = 1; i < lastColumn; i++)
            {
                if (worksheet.Cells[6, i].Value == null || string.Equals(worksheet.Cells[6, i].Value.ToString(),
                        "grand total", System.StringComparison.OrdinalIgnoreCase)) continue;
                var header = new Header { Column = i, Name = worksheet.Cells[6, i].Value.ToString() };
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

        public bool AreHeadersTranslated => NotTranslatedHeaders.Count == 0;

        public void TranslateHeaders()
        {
            var serialization = new TranslationSerialization();

            _translatedHeaders.Clear();
            _translatedHeaders = Headers.Select(header => new Header()
            {
                Column = header.Column,
                Name = header.Name
            }).ToList();

            var translations = serialization.DeserializeTranslations();
            var untranslatedAbsences = new List<Day>();

            foreach (var absence in Employees.SelectMany(d => d.Days).Where(d => d.WorkType == WorkType.Absence && string.IsNullOrWhiteSpace(d.TranslatedAbsence)))
            {
                if (untranslatedAbsences.Any(d => d.Absence == absence.Absence)) continue;
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
