using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VF_Raporty_Godzin_Pracy
{
    class Raport
    {
        public List<Pracowik> Pracownicy;
        public List<Naglowek> Naglowki;

        public Raport()
        {
            var naglowki = new List<Naglowek>();
            var pracownicy = new List<Pracowik>();
        }
    }
}
