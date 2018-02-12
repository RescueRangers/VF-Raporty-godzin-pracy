using System.Collections.Generic;
using VF_Raporty_Godzin_Pracy;

namespace WinGUI.Utility
{
    public class WyslijDoTlumaczenia
    {
        public List<Tlumaczenie> NaglowkiDoTlumaczenia { get; set; }

        public WyslijDoTlumaczenia()
        {
            NaglowkiDoTlumaczenia = new List<Tlumaczenie>();
        }
    }
}
