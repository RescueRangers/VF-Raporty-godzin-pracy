using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VF_Raporty_Godzin_Pracy
{
    [Serializable]
    public class PrzetlumaczoneNaglowki
    {
        public ObservableCollection<Tlumaczenie> ListaTlumaczen { get; set; }

        public PrzetlumaczoneNaglowki()
        {
              ListaTlumaczen = new ObservableCollection<Tlumaczenie>();
        }

        public void UstawTlumaczenia(ObservableCollection<Tlumaczenie> listaTlumaczen)
        {
            if (listaTlumaczen != null)
            {
                ListaTlumaczen = listaTlumaczen;
            }
        }

    }
}
