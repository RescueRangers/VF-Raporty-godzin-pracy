using System.Collections.Generic;
using DAL;

namespace WinGUI.Utility
{
    public class WyslijDoTlumaczenia
    {
        public List<Translation> NaglowkiDoTlumaczenia { get; set; }

        public WyslijDoTlumaczenia()
        {
            NaglowkiDoTlumaczenia = new List<Translation>();
        }
    }
}
