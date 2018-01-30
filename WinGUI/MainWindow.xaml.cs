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
            TlumaczeniaLista.ItemsSource = Tlumacz.LadujTlumaczenia();
            Environment.CurrentDirectory = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
        }

        /// <summary>
        /// Otwieramy plik excel, jezeli jest to plik xls to przerabiamy go na xlsx i tworzy z tego pliku raport.
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

            if (string.IsNullOrWhiteSpace(plikDoRaportu))
                return;

            if (plikDoRaportu.ToLower()[plikDoRaportu.Count() - 1] == 's')
            {
                plikDoRaportu = KonwertujPlikExcel.XlsDoXlsx(plikDoRaportu);
            }

            raport = UtworzRaport.Stworz(plikDoRaportu) ?? null;

            if (raport == null)
            {
                MessageBox.Show("Nie udało się stworzyć raportu.\nSprawdz plik excel "+plikDoRaportu,"Błąd podczas otwierania raportu.",MessageBoxButton.OK,MessageBoxImage.Error);
                return;
            }

            if (raport.CzyPrzetlumaczoneNaglowki() == false)
            {
                NieTlumaczone.ItemsSource = raport.PobierzNiePrzetlumaczoneNaglowki();
                Grid.SetRowSpan(TlumaczeniaLista, 1);
                LabelNieTlumaczone.Visibility = Visibility.Visible;
                NieTlumaczone.Visibility = Visibility.Visible;
                TlumaczNaglowki.Visibility = Visibility.Visible;
                MessageBox.Show("W raporcie znajdują się nieprzetłumaczone nagłówki", "Nieprzetłumaczone nagłówki.", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            Execute.IsEnabled = true;
            JedenPracownik.IsEnabled = true;
            WyborPracownika.ItemsSource = raport.PobierzPracownikowDoWidoku();
            WyborPracownika.Columns[0].SortDirection = System.ComponentModel.ListSortDirection.Descending;
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
                MessageBox.Show(ZapiszExcel.ZapiszDoExcel(raport, wybraniPracownicy, folderDoZapisu));
            }
            else
            {
                MessageBox.Show(ZapiszExcel.ZapiszDoExcel(raport, folderDoZapisu));
            }
        }

        private void JedenPracownik_Checked(object sender, RoutedEventArgs e)
        { }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void EdytujTlumaczenia_Click(object sender, RoutedEventArgs e)
        {
            if (TlumaczeniaLista.SelectedItems.Count > 0)
            {
                var naglowkiDictionary = PobierzNaglowki();
                var przetlumaczoneNaglowki = new Dictionary<string, string>();
                foreach (var tlumaczenie in naglowkiDictionary)
                {
                    var dialogTlumaczenia = new Tlumaczenia(tlumaczenie.Key, tlumaczenie.Value);
                    var wynik = dialogTlumaczenia.ShowDialog();
                    if (wynik.HasValue && wynik.Value)
                    {
                        przetlumaczoneNaglowki.Add(tlumaczenie.Key, dialogTlumaczenia.Przetlumaczone);
                    }
                }
                Tlumacz.EdytujTlumaczenia(przetlumaczoneNaglowki);
                raport.TlumaczNaglowki();
                TlumaczeniaLista.ItemsSource = null;
                TlumaczeniaLista.ItemsSource = Tlumacz.LadujTlumaczenia();
            }
            else
            {
                MessageBox.Show("Nie wybrano tłumaczeń do edycji.", "Tłumaczenia", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void UsunTlumaczenia_Click(object sender, RoutedEventArgs e)
        {
            if (TlumaczeniaLista.SelectedItems.Count > 0)
            {
                Tlumacz.UsunTlumaczenia(PobierzNaglowki());
                TlumaczeniaLista.ItemsSource = null;
                TlumaczeniaLista.ItemsSource = Tlumacz.LadujTlumaczenia();
            }
            else
            {
                MessageBox.Show("Nie wybrano tłumaczeń do edycji.","Tłumaczenia",MessageBoxButton.OK,MessageBoxImage.Information);
            }
            
        }

        private Dictionary<string, string> PobierzNaglowki()
        {
            var naglowkiDoEdycji = TlumaczeniaLista.SelectedItems.OfType<KeyValuePair<string, string>>();
            var naglowkiDictionary = new Dictionary<string, string>();
            foreach (var naglowek in naglowkiDoEdycji)
            {
                naglowkiDictionary.Add(naglowek.Key, naglowek.Value);
            }
            return naglowkiDictionary;
        }

        private void TlumaczNaglowki_Click(object sender, RoutedEventArgs e)
        {
            var przetlumaczoneNaglowki = new Dictionary<string, string>();

            foreach (Naglowek item in NieTlumaczone.SelectedItems)
            {
                //var naglowek = (Naglowek)item;
                var dialogTlumaczenia = new Tlumaczenia(item.Nazwa);
                var wynik = dialogTlumaczenia.ShowDialog();
                if (wynik.HasValue && wynik.Value)
                {
                    przetlumaczoneNaglowki.Add(item.Nazwa, dialogTlumaczenia.Przetlumaczone);
                }
            }

            Tlumacz.DodajTlumaczenia(przetlumaczoneNaglowki);

            raport.CzyscListeNieprzetlumaczonych();
            raport.TlumaczNaglowki();

            TlumaczeniaLista.ItemsSource = null;
            TlumaczeniaLista.ItemsSource = Tlumacz.LadujTlumaczenia();

            if (raport.CzyPrzetlumaczoneNaglowki() == false)
            {
                NieTlumaczone.ItemsSource = raport.PobierzNiePrzetlumaczoneNaglowki();
                Grid.SetRowSpan(TlumaczeniaLista, 1);
                LabelNieTlumaczone.Visibility = Visibility.Visible;
                NieTlumaczone.Visibility = Visibility.Visible;
                TlumaczNaglowki.Visibility = Visibility.Visible;
            }
            else
            {
                //NieTlumaczone.ItemsSource = raport.PobierzNiePrzetlumaczoneNaglowki();
                Grid.SetRowSpan(TlumaczeniaLista, 4);
                LabelNieTlumaczone.Visibility = Visibility.Hidden;
                NieTlumaczone.Visibility = Visibility.Hidden;
                TlumaczNaglowki.Visibility = Visibility.Hidden;
            }
        }
    }
}