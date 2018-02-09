using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using GalaSoft.MvvmLight.Messaging;
using VF_Raporty_Godzin_Pracy;
using VF_Raporty_Godzin_Pracy.Annotations;
using WinGUI.Utility;

namespace WinGUI.ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        #region Atrybuty
        private ObservableCollection<Naglowek> _listaNietlumaczonychNaglowkow;
        private Raport _raport;
        private readonly string _sciezkaDoXml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                                                @"\Vest-Fiber\Raporty\Tlumaczenia.xml";
        private readonly SerializacjaTlumaczen _serializacja = new SerializacjaTlumaczen();
        private PrzetlumaczoneNaglowki _przetlumaczoneNaglowki;
        private bool _wybraniPracownicyZaznaczony;
        private readonly string _folderAplikacji = AppDomain.CurrentDomain.BaseDirectory;
        private IList _wybraniPracownicy = new ArrayList();
        private IList _wybraneTlumaczenia = new ArrayList();
        private IWiadomosc Wiadomosci { get; }
        private IWyborPliku WyborPliku { get; }

        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand OtworzPlik { get; set; }
        public  ICommand ZapiszPlik { get; set; }
        public ICommand ZamknijAplikacje { get; set; }
        public ICommand UsunTlumaczenia { get; set; }
        public ICommand WyslijDoTlumaczenia { get; set; }
       

        public IList WybraneTlumaczenia
        {
            get => _wybraneTlumaczenia;
            set
            {
                _wybraneTlumaczenia = value; 
                OnPropertyChanged(nameof(WybraneTlumaczenia));
            }
        }

        public IList WybraniPracownicy
        {
            get => _wybraniPracownicy;
            set
            {
                _wybraniPracownicy = value; 
                OnPropertyChanged(nameof(WybraniPracownicy));
            }
        }

        public bool WybraniPracownicyZaznaczony
        {
            get => _wybraniPracownicyZaznaczony;
            set
            {
                _wybraniPracownicyZaznaczony = value;
                OnPropertyChanged(nameof(WybraniPracownicyZaznaczony));
            }
        }

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
        #endregion

        #region Eventy

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ZamykanieOkna(object sender, CancelEventArgs e)
        {
            _serializacja.SerializujTlumaczenia(PrzetlumaczoneNaglowki);
        }
        
        #endregion
        

        public MainWindowViewModel()
        {
            Application.Current.MainWindow.Closing += ZamykanieOkna;
            ListaNietlumaczonychNaglowkow = new ObservableCollection<Naglowek>();
            PrzetlumaczoneNaglowki = new PrzetlumaczoneNaglowki();
            Wiadomosci = new WiadomoscGui();
            WyborPliku = new WyborPlikuGui();
            LadujDane();
            LadujKomendy();
        }

        private void LadujKomendy()
        {
            OtworzPlik = new CustomCommands(OtworzXls, MozeOtworzycXls);
            ZamknijAplikacje = new CustomCommands(Zamknij, MozeZamknac);
            ZapiszPlik = new CustomCommands(ZapiszRaport, MozeZapisac);
            UsunTlumaczenia = new CustomCommands(UsunPrzetlumaczone, MozeUsunac);
            WyslijDoTlumaczenia = new CustomCommands(TlumaczNaglowki, MozeTlumaczyc);
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

        private object DoTlumaczenia()
        {
            var wiadomosc = new WyslijDoTlumaczenia() {NaglowkiDoTlumaczenia = ListaNietlumaczonychNaglowkow.ToList()};
            Messenger.Default.Send<WyslijDoTlumaczenia>(wiadomosc);
            return null;
        }

        #region Komendy
        
        private bool MozeTlumaczyc(object obj)
        {
            return ListaNietlumaczonychNaglowkow.Any();
        }

        private void TlumaczNaglowki(object obj)
        {
            DoTlumaczenia();
        }

        private bool MozeUsunac(object obj)
        {
            return WybraneTlumaczenia != null && WybraneTlumaczenia.OfType<Tlumaczenie>().Any();
        }

        private void UsunPrzetlumaczone(object obj)
        {
            var listaTlumaczen = WybraneTlumaczenia.OfType<Tlumaczenie>().ToList();
            foreach (var tlumaczenie in listaTlumaczen)
            {
                PrzetlumaczoneNaglowki.UsunTlumaczenia(tlumaczenie);
            }
        }

        private bool MozeZapisac(object obj)
        {
            return Raport != null;
        }

        private void ZapiszRaport(object obj)
        {
            var folderDoZapisu =
                WyborPliku.OtworzFolder("Wybierz folder w którym będą zapisane raporty.", _folderAplikacji);

            if (WybraniPracownicyZaznaczony)
            {
                List<Pracowik> listaPracownikow = WybraniPracownicy.OfType<Pracowik>().ToList();
                Wiadomosci.WyslijWiadomosc(ZapiszExcel.ZapiszDoExcel(Raport, listaPracownikow, folderDoZapisu), "Operacja eksportu", TypyWiadomosci.Informacja);
            }
            else
            {
                Wiadomosci.WyslijWiadomosc(ZapiszExcel.ZapiszDoExcel(Raport, folderDoZapisu), "Operacja eksportu", TypyWiadomosci.Informacja);
            }
        }

        private static bool MozeZamknac(object obj)
        {
            return true;
        }

        private static void Zamknij(object obj)
        {
            Application.Current.MainWindow.Close();
        }

        private bool MozeOtworzycXls(object obj)
        {
            return true;
        }

        private void OtworzXls(object obj)
        {
            const string plikiExcel = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx";
            var plikDoRaportu = WyborPliku.OtworzPlik("Wybierz raport w pliku Excela", plikiExcel, _folderAplikacji);

            if (plikDoRaportu.Length == 1)
            {
                Wiadomosci.WyslijWiadomosc("Nie wybrano raportu do przetworzenia", "Raport", TypyWiadomosci.Informacja);
            }

            if (plikDoRaportu.ToLower()[plikDoRaportu.Length - 1] == 's')
            {
                plikDoRaportu = KonwertujPlikExcel.XlsDoXlsx(plikDoRaportu);
            }

            Raport = UtworzRaport.Stworz(plikDoRaportu) ?? null;

            if (Raport == null)
            {
                Wiadomosci.WyslijWiadomosc("Nie udało się stworzyć raportu.\nSprawdz plik excel "+plikDoRaportu,"Błąd podczas tworzenia raportu.", TypyWiadomosci.Blad);
                return;
            }

            if (!Raport.CzyPrzetlumaczoneNaglowki())
            {
                ListaNietlumaczonychNaglowkow = Raport.PobierzNiePrzetlumaczoneNaglowki();
            }
        }
        #endregion

    }
}
