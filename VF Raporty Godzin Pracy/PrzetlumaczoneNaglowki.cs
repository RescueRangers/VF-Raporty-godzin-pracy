using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

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

        public void UsunTlumaczenia(Tlumaczenie tlumaczenie)
        {
            ListaTlumaczen.Remove(tlumaczenie);

            //foreach (var tlumaczenie in tlumaczenia)
            //{
            //    ListaTlumaczen.Remove(tlumaczenie);
            //}
        }

        public void DodajTlumaczenia(List<Tlumaczenie> tlumaczenia)
        {
            
        }
    }
}
