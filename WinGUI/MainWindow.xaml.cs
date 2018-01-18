using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;
using System;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.ObjectModel;
using System.Windows.Controls;

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

        /// <summary>
        /// Otwieramy plik esxcel, jezeli jest to plik xls to przerabiamy go na xlsx i tworzy z tego pliku raport.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                Grid.SetRowSpan(Tlumaczenia, 1);
                LabelNieTlumaczone.Visibility = Visibility.Visible;
                NieTlumaczone.Visibility = Visibility.Visible;
            }
            Execute.IsEnabled = true;
            JedenPracownik.IsEnabled = true;
            WyborPracownika.ItemsSource = raport.PobierzPracownikowDoWidoku();
            WyborPracownika.Columns[0].SortDirection = System.ComponentModel.ListSortDirection.Descending;
            foreach (var column in WyborPracownika.Columns)
            {
                column.Width = WyborPracownika.Width/2;
            }
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