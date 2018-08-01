using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using VF_Raporty_Godzin_Pracy;
using VF_Raporty_Godzin_Pracy.Annotations;
using VF_Raporty_Godzin_Pracy.Interfaces;
using WinGUI.Extensions;
using WinGUI.Servicess;
using WinGUI.Utility;

namespace WinGUI.ViewModel
{
    public sealed class MainWindowViewModel : INotifyPropertyChanged
    {
        #region Atrybuty

        #region Privates

        private readonly IProgressDialogService _progressDialog;
        private ProgressDialogOptions _option;

        private string _folderDoZapisu;
        private string _plikExcel;
        private ObservableCollection<Tlumaczenie> _listaNietlumaczonychNaglowkow;
        private Raport _raport;
        private readonly string _sciezkaDoXml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                                                @"\Vest-Fiber\Raporty\Tlumaczenia.xml";

        private const string PlikiExcel = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx";

        private readonly SerializacjaTlumaczen _serializacja = new SerializacjaTlumaczen();
        private ObservableCollection<Tlumaczenie> _przetlumaczoneNaglowki;
        private bool _wybraniPracownicyZaznaczony;
        private readonly string _myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private IList _wybraniPracownicy = new ArrayList();
        private IList _wybraneTlumaczenia = new ArrayList();
        private IWiadomosc Wiadomosci { get; }
        private IWyborPliku WyborPliku { get; }
        private IZapiszExcel ZapiszRaportDoExcel { get; }

        #endregion

        #region Publics

        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand OtworzPlik { get; set; }
        public ICommand ZapiszPlik { get; set; }
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

        public ObservableCollection<Tlumaczenie> ListaNietlumaczonychNaglowkow
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

        public ObservableCollection<Tlumaczenie> PrzetlumaczoneNaglowki
        {
            get => _przetlumaczoneNaglowki;
            set
            {
                _przetlumaczoneNaglowki = value; 
                OnPropertyChanged(nameof(PrzetlumaczoneNaglowki));
            }
        }

        #endregion

        #endregion

        #region Eventy

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ZamykanieOkna(object sender, CancelEventArgs e)
        {
            _serializacja.SerializujTlumaczenia(PrzetlumaczoneNaglowki.ToList());
        }
        
        #endregion
        
        public MainWindowViewModel(IProgressDialogService progressDialog)
        {
            _progressDialog = progressDialog;
            if (Application.Current.MainWindow != null) Application.Current.MainWindow.Closing += ZamykanieOkna;
            ListaNietlumaczonychNaglowkow = new ObservableCollection<Tlumaczenie>();
            PrzetlumaczoneNaglowki = new ObservableCollection<Tlumaczenie>();
            Wiadomosci = new WiadomoscGui();
            WyborPliku = new WyborPlikuGui();
            ZapiszRaportDoExcel = new ZapiszExcelPionowo();
            LadujDane();
            LadujKomendy();
            
        }

        private void LadujKomendy()
        {
            OtworzPlik = new CustomCommands(OtworzXlsCommand, p => true);
            ZamknijAplikacje = new CustomCommands(Zamknij, p => true);
            ZapiszPlik = new CustomCommands(ZapiszRaportCommand, MozeZapisac);
            UsunTlumaczenia = new CustomCommands(UsunPrzetlumaczone, MozeUsunac);
            WyslijDoTlumaczenia = new CustomCommands(TlumaczNaglowki, MozeTlumaczyc);
        }

        private void LadujDane()
        {
            if (!File.Exists(_sciezkaDoXml) || new FileInfo(_sciezkaDoXml).Length == 0)
            {
                const string tlumaczeniaXml =
                    "<?xml version=\"1.0\"?>\r\n<ArrayOfTlumaczenie " + 
                    "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " + 
                    "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" />";
                File.WriteAllText(_sciezkaDoXml, tlumaczeniaXml);
            }

            PrzetlumaczoneNaglowki = _serializacja.DeserializujTlumaczenia().ToObservableCollection();
        }

        #region Komendy
        
        private void OtworzXlsCommand(object obj)
        {
            _plikExcel = WyborPliku.OtworzPlik("Wybierz raport w pliku Excela", PlikiExcel, _myDocuments);

            if (string.IsNullOrWhiteSpace(_plikExcel))
            {
                Wiadomosci.WyslijWiadomosc("Nie wybrano raportu do przetworzenia", "Raport", TypyWiadomosci.Informacja);
                return;
            }

            _option = new ProgressDialogOptions
            {
                Label = "Obecna operacja: ",
                WindowTitle = "Otwieranie raportu w pliku Excel"
            };
            _progressDialog.Execute(OtworzXls, _option);

        }

        private void ZapiszRaportCommand(object obj)
        {
            _folderDoZapisu =
                WyborPliku.OtworzFolder("Wybierz folder w którym będą zapisane raporty.", _myDocuments);
            if (string.IsNullOrWhiteSpace(_folderDoZapisu))
            {
                Wiadomosci.WyslijWiadomosc("Nie wybrano folderu do zapisu", "Wybór folderu.",TypyWiadomosci.Informacja);
                return;
            }

            _option = new ProgressDialogOptions {Label = "Obecnie przetwarzany:", WindowTitle = "Zapisywanie raportów"};

            _progressDialog.Execute(ZapiszRaport, _option);
        }

        private bool MozeTlumaczyc(object obj)
        {
            return ListaNietlumaczonychNaglowkow.Any();
        }

        private void TlumaczNaglowki(object obj)
        {

            var przetlumaczone = ListaNietlumaczonychNaglowkow.Where(n => !string.IsNullOrWhiteSpace(n.Przetlumaczone)).ToList();

            if (przetlumaczone.Any())
            {
                foreach (var tlumaczenie in przetlumaczone)
                {
                    PrzetlumaczoneNaglowki.Add(tlumaczenie);
                    ListaNietlumaczonychNaglowkow.Remove(tlumaczenie);
                }

                _serializacja.SerializujTlumaczenia(PrzetlumaczoneNaglowki.ToList());
                Raport.TlumaczNaglowki();
            }
        }

        private bool MozeUsunac(object obj)
        {
            return WybraneTlumaczenia != null && WybraneTlumaczenia.OfType<Tlumaczenie>().Any();
        }

        private void UsunPrzetlumaczone(object obj)
        {
            var listaTlumaczen = WybraneTlumaczenia.OfType<Tlumaczenie>().ToList();

            var listaTLumaczenZRaportu = Raport?.Naglowki.Where(naglowek => listaTlumaczen.Contains(naglowek)).ToList();

            if (listaTLumaczenZRaportu != null && listaTLumaczenZRaportu.Any())
            {
                foreach (var tlumaczenie in listaTLumaczenZRaportu)
                {
                    ListaNietlumaczonychNaglowkow.Add(tlumaczenie.DoTlumaczenia());
                }
            }

            foreach (var tlumaczenie in listaTlumaczen)
            {
                PrzetlumaczoneNaglowki.Remove(tlumaczenie);
            }
            _serializacja.SerializujTlumaczenia(PrzetlumaczoneNaglowki.ToList());
        }

        private bool MozeZapisac(object obj)
        {
            return Raport != null;
        }

        private async void ZapiszRaport(CancellationToken cancellationToken, IProgress<ProgressReport> progress)
        {
            var result = "";
            var currentPracownik = 0;
            int maxPracownik;

            var progressReport = new ProgressReport();

            if (WybraniPracownicyZaznaczony)
            {
                maxPracownik = WybraniPracownicy.Count + 1;
                foreach (var pracowik in WybraniPracownicy)
                {
                    var wybranyPracownik = (Pracowik)pracowik;
                    currentPracownik++;

                    progressReport.CurrentTaskNumber = currentPracownik;
                    progressReport.MaxTaskNumber = maxPracownik;
                    progressReport.IsIndeterminate = false;
                    progressReport.CurrentTask = wybranyPracownik.NazwaPracownika();

                    cancellationToken.ThrowIfCancellationRequested();
                    progress.Report(progressReport);
                    result = await ZapiszRaportDoExcel.ZapiszDoExcel(Raport, _folderDoZapisu, wybranyPracownik);
                }
                Wiadomosci.WyslijWiadomosc(result, "Operacja eksportu", TypyWiadomosci.Informacja);
            }
            else
            {
                maxPracownik = Raport.Pracownicy.Count + 1;
                foreach (var pracowik in Raport.Pracownicy)
                {
                    currentPracownik++;

                    progressReport.CurrentTaskNumber = currentPracownik;
                    progressReport.MaxTaskNumber = maxPracownik;
                    progressReport.IsIndeterminate = false;
                    progressReport.CurrentTask = pracowik.NazwaPracownika();

                    cancellationToken.ThrowIfCancellationRequested();
                    progress.Report(progressReport);
                    result = await ZapiszRaportDoExcel.ZapiszDoExcel(Raport, _folderDoZapisu, pracowik);
                }
                Wiadomosci.WyslijWiadomosc(result, "Operacja eksportu", TypyWiadomosci.Informacja);
            }
        }

        private static void Zamknij(object obj)
        {
            Debug.Assert(Application.Current.MainWindow != null, "Application.Current.MainWindow != null");
            Application.Current.MainWindow.Close();
        }

        private void OtworzXls(CancellationToken cancellationToken, IProgress<ProgressReport> progress)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var progressReport = new ProgressReport
            {
                IsIndeterminate = true
            };

            try
            {
                if (_plikExcel.ToLower()[_plikExcel.Length - 1] == 's')
                {
                    progressReport.CurrentTask = "Konwertowanie pliku do .xlsx";
                    progress.Report(progressReport);
                    _plikExcel = KonwertujPlikExcel.XlsDoXlsx(_plikExcel);
                }
            }
            catch (Exception e)
            {
                Wiadomosci.WyslijWiadomosc(e.Message, e.Source, TypyWiadomosci.Blad);
                throw;
            }

            progressReport.CurrentTask = "Tworzenie raportu";
            progress.Report(progressReport);
            Raport = UtworzRaport.Stworz(_plikExcel) ?? null;

            if (Raport == null)
            {
                Wiadomosci.WyslijWiadomosc("Nie udało się stworzyć raportu.\nSprawdz plik excel "+_plikExcel,"Błąd podczas tworzenia raportu.", TypyWiadomosci.Blad);
                return;
            }

            if (Raport.CzyPrzetlumaczoneNaglowki()) return;
            foreach (var naglowek in Raport.NiePrzetlumaczoneNaglowki)
            {
                Application.Current.Dispatcher.Invoke((Action)delegate // <--- HERE
                {
                    ListaNietlumaczonychNaglowkow.Add(naglowek.DoTlumaczenia());
                });
                //ListaNietlumaczonychNaglowkow.Add(naglowek.DoTlumaczenia());
            }
        }
        #endregion
    }
}
