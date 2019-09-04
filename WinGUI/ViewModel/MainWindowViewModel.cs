using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using DAL;
using DAL.Interfaces;
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
        private ObservableCollection<Translation> _listaNietlumaczonychNaglowkow;
        private Report _report;
        private readonly string _sciezkaDoXml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                                                @"\Vest-Fiber\Raporty\Tlumaczenia.xml";

        private const string PlikiExcel = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx";

        private readonly TranslationSerialization _serializacja = new TranslationSerialization();
        private ObservableCollection<Translation> _przetlumaczoneNaglowki;
        private bool _wybraniPracownicyZaznaczony;
        private readonly string _myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private IList _wybraniPracownicy = new ArrayList();
        private IList _wybraneTlumaczenia = new ArrayList();
        private Employee _selectedEmployee;
        private IWiadomosc Wiadomosci { get; }
        private IWyborPliku WyborPliku { get; }
        private ISaveExcel SaveRaportDoExcel { get; }

        #endregion

        #region Publics

        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand OtworzPlik { get; set; }
        public ICommand ZapiszPlik { get; set; }
        public ICommand ZamknijAplikacje { get; set; }
        public ICommand UsunTlumaczenia { get; set; }
        public ICommand WyslijDoTlumaczenia { get; set; }

        public Employee SelectedEmployee
        {
            get => _selectedEmployee;

            set
            {
                if (_selectedEmployee != value)
                {
                    _selectedEmployee = value;
                    OnPropertyChanged(nameof(SelectedEmployee));
                }
            }
        }

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

        public ObservableCollection<Translation> ListaNietlumaczonychNaglowkow
        {
            get => _listaNietlumaczonychNaglowkow;
            set
            {
                _listaNietlumaczonychNaglowkow = value; 
                OnPropertyChanged(nameof(ListaNietlumaczonychNaglowkow));
            }
        }

        public Report Report
        {
            get => _report;
            set
            {
                _report = value; 
                OnPropertyChanged(nameof(Report));
            }
        }

        public ObservableCollection<Translation> PrzetlumaczoneNaglowki
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

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ZamykanieOkna(object sender, CancelEventArgs e)
        {
            _serializacja.SerializeTranslations(PrzetlumaczoneNaglowki.ToList());
        }
        
        #endregion

        public MainWindowViewModel(IProgressDialogService progressDialog)
        {
            _progressDialog = progressDialog;
            if (Application.Current.MainWindow != null) Application.Current.MainWindow.Closing += ZamykanieOkna;
            ListaNietlumaczonychNaglowkow = new ObservableCollection<Translation>();
            PrzetlumaczoneNaglowki = new ObservableCollection<Translation>();
            Wiadomosci = new WiadomoscGui();
            WyborPliku = new WyborPlikuGui();
            SaveRaportDoExcel = new SaveExcelVertical();
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

            PrzetlumaczoneNaglowki = _serializacja.DeserializeTranslations().ToObservableCollection();
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
            var przetlumaczone = ListaNietlumaczonychNaglowkow.Where(n => !string.IsNullOrWhiteSpace(n.Translated)).ToList();

            if (przetlumaczone.Any())
            {
                foreach (var tlumaczenie in przetlumaczone)
                {
                    PrzetlumaczoneNaglowki.Add(tlumaczenie);
                    ListaNietlumaczonychNaglowkow.Remove(tlumaczenie);
                }

                _serializacja.SerializeTranslations(PrzetlumaczoneNaglowki.ToList());
                Report.TranslateHeaders();
            }
        }

        private bool MozeUsunac(object obj)
        {
            return WybraneTlumaczenia != null && WybraneTlumaczenia.OfType<Translation>().Any();
        }

        private void UsunPrzetlumaczone(object obj)
        {
            var listaTlumaczen = WybraneTlumaczenia.OfType<Translation>().ToList();

            var listaTLumaczenZRaportu = Report?.Headers.Where(naglowek => listaTlumaczen.Contains(naglowek)).ToList();

            if (listaTLumaczenZRaportu != null && listaTLumaczenZRaportu.Any())
            {
                foreach (var tlumaczenie in listaTLumaczenZRaportu)
                {
                    ListaNietlumaczonychNaglowkow.Add(tlumaczenie.ToTranslate());
                }
            }

            foreach (var tlumaczenie in listaTlumaczen)
            {
                PrzetlumaczoneNaglowki.Remove(tlumaczenie);
            }
            _serializacja.SerializeTranslations(PrzetlumaczoneNaglowki.ToList());
        }

        private bool MozeZapisac(object obj)
        {
            return Report != null;
        }

        private async void ZapiszRaport(CancellationToken cancellationToken, IProgress<ProgressReport> progress)
        {
            var result = "";
            var currentPracownik = 0;
            int maxPracownik;

            var progressReport = new ProgressReport();

            if (WybraniPracownicyZaznaczony)
            {
                maxPracownik = WybraniPracownicy.Count;
                foreach (var pracowik in WybraniPracownicy)
                {
                    var wybranyPracownik = (Employee)pracowik;
                    currentPracownik++;

                    progressReport.CurrentTaskNumber = currentPracownik;
                    progressReport.MaxTaskNumber = maxPracownik;
                    progressReport.IsIndeterminate = false;
                    progressReport.CurrentTask = wybranyPracownik.FullName;

                    cancellationToken.ThrowIfCancellationRequested();
                    progress.Report(progressReport);
                    result = await SaveRaportDoExcel.SaveExcel(Report, _folderDoZapisu, wybranyPracownik);

                    if (result != "Operacja")
                    {
                        
                    }
                }
                Wiadomosci.WyslijWiadomosc(result, "Operacja eksportu", TypyWiadomosci.Informacja);
            }
            else
            {
                maxPracownik = Report.Employees.Count;
                foreach (var pracowik in Report.Employees)
                {
                    currentPracownik++;

                    progressReport.CurrentTaskNumber = currentPracownik;
                    progressReport.MaxTaskNumber = maxPracownik;
                    progressReport.IsIndeterminate = false;
                    progressReport.CurrentTask = pracowik.FullName;

                    cancellationToken.ThrowIfCancellationRequested();
                    progress.Report(progressReport);
                    result = await SaveRaportDoExcel.SaveExcel(Report, _folderDoZapisu, pracowik);

                    if (result != Properties.Resources.Success)
                    {
                        Wiadomosci.WyslijWiadomosc(result, "Operacja eksportu", TypyWiadomosci.Blad);
                        return;
                    }

                }
                Wiadomosci.WyslijWiadomosc(result, "Operacja eksportu", TypyWiadomosci.Informacja);
            }
        }

        private static void Zamknij(object obj)
        {
            Application.Current.MainWindow?.Close();
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
                    _plikExcel = ConvertExcel.XlsToXlsx(_plikExcel);
                }
            }
            catch (Exception e)
            {
                Wiadomosci.WyslijWiadomosc("Nie udało się skonwertować pliku do xlsx, spróbuj skonvertować plik ręcznie", e.Source, TypyWiadomosci.Blad);
                throw;
            }

            progressReport.CurrentTask = "Tworzenie raportu";
            progress.Report(progressReport);
            try
            {
                Report = Report.Create(_plikExcel);
            }
            catch (FileLoadException e)
            {
                Console.WriteLine(e);
                Wiadomosci.WyslijWiadomosc("Nie udało się stworzyć raportu.\nSprawdz plik excel "+_plikExcel,"Błąd podczas tworzenia raportu.", TypyWiadomosci.Blad);
                return;
            }

            if (Report.AreHeadersTranslated) return;
            foreach (var naglowek in Report.NotTranslatedHeaders)
            {
                Application.Current.Dispatcher.Invoke((Action)delegate // <--- HERE
                {
                    ListaNietlumaczonychNaglowkow.Add(new Translation(naglowek.Absence, ""));
                });
            }
        }
        #endregion
    }
}
