using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using VF_Raporty_Godzin_Pracy;
using VF_Raporty_Godzin_Pracy.Annotations;
using WinGUI.Utility;

namespace WinGUI.ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<Naglowek> _listaNietlumaczonychNaglowkow;
        private Raport _raport;
        private readonly string _sciezkaDoXml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                                                @"\Vest-Fiber\Raporty\Tlumaczenia.xml";
        private readonly SerializacjaTlumaczen _serializacja = new SerializacjaTlumaczen();

        private PrzetlumaczoneNaglowki _przetlumaczoneNaglowki;

        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand OtworzPlik { get; set; }

        public ObservableCollection<Naglowek> ListaNietlumaczonychNaglowkow
        {
            get => _listaNietlumaczonychNaglowkow;
            set
            {
                _listaNietlumaczonychNaglowkow = value; 
                OnPropertyChanged(nameof(ListaNietlumaczonychNaglowkow));
            }
        }

        public Raport Raport
        {
            get => _raport;
            set
            {
                _raport = value; 
                OnPropertyChanged(nameof(Raport));
            }
        }

        public PrzetlumaczoneNaglowki PrzetlumaczoneNaglowki
        {
            get => _przetlumaczoneNaglowki;
            set
            {
                _przetlumaczoneNaglowki = value; 
                OnPropertyChanged(nameof(PrzetlumaczoneNaglowki));
            }
        }

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public MainWindowViewModel()
        {
            ListaNietlumaczonychNaglowkow = new ObservableCollection<Naglowek>();
            PrzetlumaczoneNaglowki = new PrzetlumaczoneNaglowki();
            LadujDane();
            LadujKomendy();
        }

        private void LadujKomendy()
        {
            OtworzPlik = new CustomCommands(OtworzXls, MozeOtworzycXls);
        }

        private bool MozeOtworzycXls(object obj)
        {
            return true;
        }

        private void OtworzXls(object obj)
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

            
        }

        private void LadujDane()
        {
            if (!File.Exists(_sciezkaDoXml) || new FileInfo(_sciezkaDoXml).Length == 0)
            {
                const string tlumaczeniaXml =
                    "<?xml version=\"1.0\"?>\r\n<PrzetlumaczoneNaglowki" 
                    + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" 
                    + "\r\n  <ListaTlumaczen />\r\n</PrzetlumaczoneNaglowki>";
                File.WriteAllText(_sciezkaDoXml,tlumaczeniaXml);
            }

            PrzetlumaczoneNaglowki.UstawTlumaczenia(_serializacja.DeserializujTlumaczenia());
        }
    }
}
