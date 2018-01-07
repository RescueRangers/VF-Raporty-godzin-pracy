using System;
using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Dzien
    {
        private List<decimal> _godziny = new List<decimal>();
        public DateTime Date;

        public List<decimal> GetGodziny()
        {
            return _godziny;
        }

        public void SetGodziny(List<decimal> godziny)
        {
            _godziny = godziny;
        }
    }
}