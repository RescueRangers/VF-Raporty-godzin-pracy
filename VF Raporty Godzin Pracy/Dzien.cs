using System;
using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Dzien
    {
        public DateTime Date;
        private List<decimal> _godziny = new List<decimal>();

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