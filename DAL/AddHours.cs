using System;
using System.Collections.Generic;
using System.Diagnostics;
using OfficeOpenXml;

namespace DAL
{
    class AddHours
    {
        public static List<decimal> GetHours(ExcelWorksheet arkusz, int indeks, List<Header> headers)
        {
            var hours = new List<decimal>();
            foreach (var header in headers)
            {
                if (header != null)
                {
                    Debug.Assert(header.Column != null, "naglowek.Kolumna != null");
                    hours.Add(Convert.ToDecimal(arkusz.Cells[indeks, (int)header.Column].Value));
                }
            }
            return hours;
        }
    }
}
