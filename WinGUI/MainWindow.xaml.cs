using System.Linq;
using System.Windows;
using VF_Raporty_Godzin_Pracy;
using System;
using System.Collections.ObjectModel;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public ObservableCollection<Naglowek> ListaNieTlumaczonychNaglowkow;
        public Raport Raport;

        public PrzetlumaczoneNaglowki Przetlumaczone = new PrzetlumaczoneNaglowki();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void TlumaczNaglowki_Click(object sender, RoutedEventArgs e)
        {
            var listaNietlumaczonych = new Naglowek[NieTlumaczone.Items.Count];
            NieTlumaczone.Items.CopyTo(listaNietlumaczonych,0);
            
            for (int i = 0; i < listaNietlumaczonych.Count(); i++)
            {
                var naglowekDoTlumaczenia = listaNietlumaczonych[i];
                var dialogTlumaczenia = new Tlumaczenia(naglowekDoTlumaczenia);
                var wynik = dialogTlumaczenia.ShowDialog();
                if (!wynik.HasValue || !wynik.Value) continue;
                Raport.DodajTlumaczenie(dialogTlumaczenia.Naglowek);
                Przetlumaczone.ListaTlumaczen.Add(dialogTlumaczenia.Naglowek);
            }
        }
    }
}