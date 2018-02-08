using System.Linq;
using System.Windows;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;
using System;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.IO;
using System.ComponentModel;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public ObservableCollection<Naglowek> ListaNieTlumaczonychNaglowkow;
        public Raport Raport;

        private readonly string _sciezkaDoXml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                                       @"\Vest-Fiber\Raporty\Tlumaczenia.xml";

        private readonly SerializacjaTlumaczen _serializacja = new SerializacjaTlumaczen();

        public PrzetlumaczoneNaglowki Przetlumaczone = new PrzetlumaczoneNaglowki();

        public MainWindow()
        {
            InitializeComponent();
            
            //Jezeli nie istnieje plik z tlumaczeniami lub plik z tlumaczeniami jest pusty tworzy szkielet pliku z tlumaczeniami
            if (!File.Exists(_sciezkaDoXml) || new FileInfo(_sciezkaDoXml).Length == 0)
            {
                const string tlumaczeniaXml =
                    "<?xml version=\"1.0\"?>\r\n<PrzetlumaczoneNaglowki" 
                    + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" 
                    + "\r\n  <ListaTlumaczen />\r\n</PrzetlumaczoneNaglowki>";
                File.WriteAllText(_sciezkaDoXml,tlumaczeniaXml);
            }

            Przetlumaczone.UstawTlumaczenia(_serializacja.DeserializujTlumaczenia());

            DataContext = Raport;

            TlumaczeniaLista.DataContext = Przetlumaczone;
        }

        /// <summary>
        /// Otwieramy plik excel, jezeli jest to plik xls to przerabiamy go na xlsx i tworzy z tego pliku raport.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            const string plikiExcel = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx";
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

            if (plikDoRaportu.ToLower()[plikDoRaportu.Length - 1] == 's')
            {
                plikDoRaportu = KonwertujPlikExcel.XlsDoXlsx(plikDoRaportu);
            }

            Raport = UtworzRaport.Stworz(plikDoRaportu) ?? null;

            if (Raport == null)
            {
                MessageBox.Show("Nie udało się stworzyć raportu.\nSprawdz plik excel "+plikDoRaportu,"Błąd podczas tworzenia raportu.",MessageBoxButton.OK,MessageBoxImage.Error);
                return;
            }

            if (Raport.CzyPrzetlumaczoneNaglowki() == false)
            {
                Grid.SetRowSpan(TlumaczeniaLista, 1);
                LabelNieTlumaczone.Visibility = Visibility.Visible;
                NieTlumaczone.Visibility = Visibility.Visible;
                NieTlumaczone.DataContext = Raport;
                TlumaczNaglowki.Visibility = Visibility.Visible;
                MessageBox.Show("W raporcie znajdują się nieprzetłumaczone nagłówki", "Nieprzetłumaczone nagłówki.", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            Execute.IsEnabled = true;
            JedenPracownik.IsEnabled = true;
            WyborPracownika.DataContext = Raport;
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
                var wybraniPracownicy = WyborPracownika.SelectedItems.Cast<Pracowik>().ToList();
                MessageBox.Show(ZapiszExcel.ZapiszDoExcel(Raport, wybraniPracownicy, folderDoZapisu), "Operacja eksportu", MessageBoxButton.OK,MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(ZapiszExcel.ZapiszDoExcel(Raport, folderDoZapisu), "Operacja eksportu", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void JedenPracownik_Checked(object sender, RoutedEventArgs e)
        { }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            //_serializacja.SerializujTlumaczenia(Przetlumaczone);
            Close();
        }

        private void UsunTlumaczenia_Click(object sender, RoutedEventArgs e)
        {
            if (TlumaczeniaLista.SelectedItems.Count > 0)
            {
                for (int i = 0; i < TlumaczeniaLista.SelectedItems.Count; i++)
                {
                    var naglowekDoUsuniecia = (Tlumaczenie) TlumaczeniaLista.SelectedItems[i];
                    Przetlumaczone.UsunTlumaczenia(naglowekDoUsuniecia);
                }
            }
            else
            {
                MessageBox.Show("Nie wybrano tłumaczeń do edycji.","Tłumaczenia",MessageBoxButton.OK,MessageBoxImage.Information);
            }
        }

        private void TlumaczNaglowki_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < NieTlumaczone.SelectedItems.Count; i++)
            {
                var naglowekDoTlumaczenia = (Naglowek) NieTlumaczone.SelectedItems[i];
                var dialogTlumaczenia = new Tlumaczenia(naglowekDoTlumaczenia);
                var wynik = dialogTlumaczenia.ShowDialog();
                if (!wynik.HasValue || !wynik.Value) continue;
                Raport.DodajTlumaczenie(dialogTlumaczenia.Naglowek);
                Przetlumaczone.ListaTlumaczen.Add(dialogTlumaczenia.Naglowek);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            _serializacja.SerializujTlumaczenia(Przetlumaczone);
        }
    }
}