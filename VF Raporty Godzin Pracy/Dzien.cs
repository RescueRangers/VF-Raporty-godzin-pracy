using System;
using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Dzien
    {
        public DateTime Date;
        private List<Godzina> _godziny = new List<Godzina>();

        public List<Godzina> GetGodziny()
        {
            return _godziny;
        }

        public void SetGodziny(List<Godzina> godziny)
        {
            _godziny = godziny;
        }
    }
}