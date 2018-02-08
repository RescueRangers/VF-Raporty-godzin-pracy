using System.Windows;
using VF_Raporty_Godzin_Pracy;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for Tlumaczenia.xaml
    /// </summary>
    public partial class Tlumaczenia : Window
    {
        public string Przetlumaczone { get { return DoPrzetlumaczenia.Text; } }

        public Tlumaczenie Naglowek { get; set; }

        public Tlumaczenia(Naglowek naglowek)
        {
            InitializeComponent();
            Naglowek = new Tlumaczenie
            {
                Nazwa = naglowek.Nazwa,
                Kolumna = naglowek.Kolumna,
                Przetlumaczone = ""
            };
            DataContext = Naglowek;
        }

        //public Tlumaczenia(string oryginal)
        //{
        //    InitializeComponent();
        //    Oryginal.Content = oryginal;
        //}

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
