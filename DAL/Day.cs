using System;
using System.Collections.Generic;

namespace DAL
{
    public class Day
    {
        public DateTime Date;

        public List<decimal> Hours { get; set; } = new List<decimal>();

        public void SetHours(List<decimal> godziny)
        {
            Hours = godziny;
        }
    }
}