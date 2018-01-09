using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;
using System;
using Microsoft.WindowsAPICodePack.Dialogs;

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
            JedenPracownik.IsEnabled = true;
            WyborPracownika.ItemsSource = raport.GetNazwyPracownikow();
        }

        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            var folderDoZapisu = "";
            var wyborFolderu = new CommonOpenFileDialog
            {
                Title = "Wybierz folder w którym będą zapisane rporty.",
                IsFolderPicker = true,
                InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
            };

            if (wyborFolderu.ShowDialog() == CommonFileDialogResult.Ok)
            {
                folderDoZapisu = wyborFolderu.FileName;
            }

            if (JedenPracownik.IsChecked == true)
            {
                var wybraniPracownicy = new List<string>();
                foreach (var item in WyborPracownika.SelectedItems)
                {
                    wybraniPracownicy.Add(item.ToString());
                }
                ZapiszExcel.ZapiszDoExcel(raport, wybraniPracownicy);
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

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
