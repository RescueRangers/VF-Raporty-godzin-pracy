using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VF_Raporty_Godzin_Pracy;

namespace WinGUI.Utility
{
    public class WyslijDoTlumaczenia
    {
        public List<Naglowek> NaglowkiDoTlumaczenia { get; set; }

        public WyslijDoTlumaczenia()
        {
            NaglowkiDoTlumaczenia = new List<Naglowek>();
        }
    }
}
