using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Raport raport;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var plikiExcel = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx";
                var plikDoRaportu = "";
                var otworzPlik = new OpenFileDialog
                {
                    Filter = plikiExcel
                };
                if (otworzPlik.ShowDialog() == true)
                {
                    plikDoRaportu = otworzPlik.FileName;
                }
                if (plikDoRaportu.ToLower()[plikDoRaportu.Count() - 1] == 's')
                {
                    plikDoRaportu = KonwertujPlikExcel.XlsDoXlsx(plikDoRaportu);
                }
                raport = UtworzRaport.Stworz(plikDoRaportu);
                Execute.IsEnabled = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Nie udało się otworzyć raportu.");
            }
        }

        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            if (JedenPracownik.IsChecked == true)
            {
                ZapiszExcel.ZapiszDoExcel(raport, WyborPracownika.SelectedIndex);
            }
            else
            {
                ZapiszExcel.ZapiszDoExcel(raport);
            }
        }

        private void JedenPracownik_Checked(object sender, RoutedEventArgs e)
        {
            WyborPracownika.IsEnabled = true;
        }
    }
}
