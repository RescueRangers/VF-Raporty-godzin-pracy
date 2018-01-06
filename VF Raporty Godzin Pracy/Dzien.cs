using System;
using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Dzien
    {
        public DateTime Date;
        public List<Godzina> Godziny;

        public Dzien()
        {
            Godziny = new List<Godzina>();
        }
    }
}