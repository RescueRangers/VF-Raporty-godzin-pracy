using System.Collections.Generic;
using DAL;

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
