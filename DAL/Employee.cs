using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml;

namespace DAL
{
    public class Employee
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }

        public List<Day> Days { get; set; } = new List<Day>();

        private int _startIndex;
        private int _endIndex;

        public List<Day> GetDays()
        {
            return Days;
        }

        public void FillDays(ExcelWorksheet worksheet, List<Header> headers)
        {
            var days = new List<Day>();
            for (var i = _startIndex; i < _endIndex; i++)
            {
                var day = new Day
                {
                    Date = DateTime.ParseExact(worksheet.Cells[i, 7].Text,"dd-MM-yyyy",new CultureInfo("pl-PL"))
                };
                day.SetHours(worksheet, i, headers);
                days.Add(day);
            }

            Days = days;
        }

        /// <summary>
        /// Zwraca string z nazwiskiem i imieniem pracownika
        /// </summary>
        /// <returns></returns>
        public string FullName
        {
            get { return $"{ LastName} { FirstName}";
        }
    }

        /// <summary>
        /// Zwracy daty
        /// </summary>
        /// <returns></returns>
        public List<DateTime> GetDates()
        {
            if (Days.Count == 0) throw new InvalidOperationException("Lista dni jest pusta");
            var dates = new List<DateTime>();
            foreach (var day in Days)
            {
                dates.Add(day.Date);
            }
            return dates;
        }

        public void SetStartIndex(int startIndeks)
        {
            _startIndex = startIndeks;
        }

        public void SetEndIndex(int koniecIndeks)
        {
            _endIndex = koniecIndeks;
        }
    }
}