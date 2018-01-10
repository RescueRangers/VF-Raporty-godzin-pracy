using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;
using System;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.ObjectModel;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<Naglowek> ListaNieTlumaczonychNaglowkow;
        private Raport raport;

        public MainWindow()
        {
            InitializeComponent();
            Tlumaczenia.ItemsSource = Tlumacz.LadujTlumaczenia();
        }

        private void Open_Click(object sender, RoutedEventArgs e)
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

            if (raport.CzyPrzetlumaczoneNaglowki() == false)
            {
                NieTlumaczone.ItemsSource = raport.PobierzNiePrzetlumaczoneNaglowki();
                NieTlumaczone.Visibility = Visibility.Visible;
            }
            Execute.IsEnabled = true;
            JedenPracownik.IsEnabled = true;
            WyborPracownika.ItemsSource = raport.PobierzPracownikowDoWidoku();
            WyborPracownika.Columns[2].Visibility = Visibility.Hidden;
            WyborPracownika.Columns[3].Visibility = Visibility.Hidden;
            WyborPracownika.Columns[4].Visibility = Visibility.Hidden;
        }

        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            var folderDoZapisu = "";
            var wyborFolderu = new CommonOpenFileDialog
            {
                Title = "Wybierz folder w którym będą zapisane raporty.",
                IsFolderPicker = true,
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            if (wyborFolderu.ShowDialog() == CommonFileDialogResult.Ok)
            {
                folderDoZapisu = wyborFolderu.FileName;
            }

            if (JedenPracownik.IsChecked == true)
            {
                var wybraniPracownicy = new List<Pracowik>();
                foreach (Pracowik item in WyborPracownika.SelectedItems)
                {
                    wybraniPracownicy.Add(item);
                }
                ZapiszExcel.ZapiszDoExcel(raport, wybraniPracownicy, folderDoZapisu);
            }
            else
            {
                ZapiszExcel.ZapiszDoExcel(raport, folderDoZapisu);
            }
        }

        private void JedenPracownik_Checked(object sender, RoutedEventArgs e)
        {
            WyborPracownika.IsEnabled = true;
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
