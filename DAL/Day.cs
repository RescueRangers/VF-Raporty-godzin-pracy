using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace DAL
{
    public class Day
    {
        private const string WorkHoursName = "NORMALPLN";
        private const string Overtime50Name = "NADGODINY1";
        private const string Overtime100Name = "NADGODZINY2";

        public DateTime Date { get; set; }

        public decimal? WorkHour { get; set; }
        public decimal? Overtime50 { get; set; }
        public decimal? Overtime100 { get; set; }
        public string Absence { get; set; }
        public string TranslatedAbsence { get; set; }

        public WorkType WorkType { get; set; }

        public void SetHours(ExcelWorksheet worksheet, int index, List<Header> headers)
        {
            var workHoursIndex = headers.Find(h => h.Name == WorkHoursName)?.Column;
            var overtime50Index = headers.Find(h => h.Name == Overtime50Name)?.Column;
            var overtime100Index = headers.Find(h => h.Name == Overtime100Name)?.Column;

            if (workHoursIndex.HasValue)
            {
                var test = decimal.TryParse(worksheet.Cells[index, workHoursIndex.Value].Text, out var value);

                if(test)
                {
                    WorkHour = value;
                    WorkType = WorkType.Normal;
                }
            }
            if (overtime50Index.HasValue)
            {
                var test = decimal.TryParse(worksheet.Cells[index, overtime50Index.Value].Text, out var value);

                if(test)
                {
                    Overtime50 = value;
                    WorkType = WorkType.Overtime1;
                }
            }
            if (overtime100Index.HasValue)
            {
                var test = decimal.TryParse(worksheet.Cells[index, overtime100Index.Value].Text, out var value);

                if(test)
                {
                    Overtime100 = value;
                    WorkType = WorkType == WorkType.Overtime1 ? WorkType.Overtimes : WorkType.Overtime2;
                }
            }
            if(!WorkHour.HasValue && !Overtime50.HasValue && !Overtime100.HasValue)
            {
                var row = worksheet.Cells[index, headers[0].Column.Value, index, headers.Last().Column.Value];
                var test = row.First(cell => decimal.TryParse(cell.Text, out var result) && result > 0);
                Absence = headers.Find(h => h.Column == test.End.Column).Name;
                WorkType = WorkType.Absence;
            }
        }
    }
}