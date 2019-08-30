using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OfficeOpenXml;

namespace DAL
{
    public class Day
    {
        private const string WorkHoursName = "NORMALPLN";
        private const string Overtime50Name = "NADGODINY1";
        private const string Overtime100Name = "NADGODINY2";

        public DateTime Date;

        //public List<decimal> Hours { get; set; } = new List<decimal>();

        public decimal WorkHour { get; set; }
        public decimal Overtime50 { get; set; }
        public decimal Overtime100 { get; set; }
        public string Absence { get; set; }
        public string TranslatedAbsence { get; set; }

        public WorkType WorkType { get; set; }

        public void SetHours(ExcelWorksheet worksheet, int index, List<Header> headers)
        {
            //var hours = new List<decimal>();

            var workHoursIndex = headers.Find(h => h.Name == WorkHoursName).Column;
            var overtime50Index = headers.Find(h => h.Name == Overtime50Name).Column;
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
                    WorkType = overtime50Index.HasValue ? WorkType.Overtimes : WorkType.Overtime2;
                }
            }
            if(!workHoursIndex.HasValue && !overtime50Index.HasValue && !overtime100Index.HasValue)
            {

                Absence = headers.Find(h => h.Column.HasValue && h.Column.Value > 0).Name;
                WorkType = WorkType.Absence;
            }
        }
    }
}