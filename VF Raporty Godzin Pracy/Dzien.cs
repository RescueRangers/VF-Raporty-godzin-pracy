using System;
using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Dzien
    {
        public DateTime Date;
        private List<decimal> _godziny = new List<decimal>();

        public List<decimal> Godziny
        {
            get => _godziny;
            set => _godziny = value;
        }

        public void SetGodziny(List<decimal> godziny)
        {
            _godziny = godziny;
        }
  
    }
}