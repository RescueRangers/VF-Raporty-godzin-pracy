using System;
using System.Collections.Generic;

namespace DAL
{
    public class Dzien
    {
        /// <summary>
        /// Lista godzin w danym dniu, 
        /// pozycja godziny odpowiada naglowkowi 
        /// z listy naglowkow
        /// </summary>
        private List<decimal> _godziny = new List<decimal>();
        public DateTime Date;

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