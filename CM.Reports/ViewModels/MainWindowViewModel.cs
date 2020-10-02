using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Caliburn.Micro;
using CM.Reports.Properties;
using CM.Reports.Utility;
using DAL;
using MahApps.Metro.Controls.Dialogs;

namespace CM.Reports.ViewModels
{
    internal class MainWindowViewModel : Conductor<ReportViewModel>.Collection.AllActive
    {
        private IWindowManager _windowManager;
        private TranslationSerialization _translationSerialization;
        private readonly IDialogCoordinator _dialogCoordinator;
        private readonly IIODialogs _ioDialogs;
        private ReportViewModel _report;
        private bool _isBusy = false;
        private int _currentEmployee = 1;
        private ProgressDialogController _progressDialogController;
        private string _currentEmployeeName;

        private int CurrentEmployee
        {
            set
            {
                if (value == _currentEmployee) return;
                _currentEmployee = value;
                _progressDialogController.SetProgress(value);
            }
        }

        private string CurrentEmployeeName
        {
            set
            {
                if (value == _currentEmployeeName) return;
                _currentEmployeeName = value;
                _progressDialogController.SetMessage(value);
            }
        }

        public ReportViewModel Report
        {
            get => _report;
            set
            {
                if (Equals(value, _report)) return;
                _report = value;
                NotifyOfPropertyChange(() => Report);
            }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (value == _isBusy) return;
                _isBusy = value;
                NotifyOfPropertyChange(() => IsBusy);
            }
        }

        public MainWindowViewModel(IWindowManager windowManager, IIODialogs ioDialogs, ReportViewModel report, IDialogCoordinator dialogCoordinator, TranslationSerialization translationSerialization)
        {
            _windowManager = windowManager;
            _ioDialogs = ioDialogs;
            _report = report;
            Items.Add(_report);
            _dialogCoordinator = dialogCoordinator;
            _translationSerialization = translationSerialization;
        }

        public async Task OpenExcelReport()
        {
            var initialDirectory = string.IsNullOrWhiteSpace(Settings.Default.InitialOpenDirectory) ? AppDomain.CurrentDomain.BaseDirectory : Settings.Default.InitialOpenDirectory;

            var filePath = _ioDialogs.OpenFile("Otwórz plik z eksportem godzin pracy.",
                initialDirectory);

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                await GenerateReport(filePath);
            }
        }

        private async Task GenerateReport(string filePath)
        {
            IsBusy = true;

            //var report = DAL.Report.Create(filePath);
            var report = await Task.Run(() => DAL.Report.Create(filePath, _translationSerialization));
            _report.MapData(report);
            Settings.Default.InitialOpenDirectory = new FileInfo(filePath).DirectoryName;
            Settings.Default.Save();

            IsBusy = false;
        }

        public async Task ReportDrop(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var file = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (file[0].EndsWith(".xls") || file[0].EndsWith(".xlsx"))
                {
                    await GenerateReport(file[0]);
                }
            }
        }

        public void OpenTranslations()
        {
            var translations = new TranslationsViewModel(_translationSerialization);

            if (Report.NotTranslatedHeaders != null) translations.HeadersToTranslate = new ObservableCollection<Translation>(Report.NotTranslatedHeaders);

            _windowManager.ShowWindow(translations);
        }

        public async Task ExportToExcel()
        {
            var p = new Progress<Tuple<int, string>>();
            p.ProgressChanged += (_, i) =>
            {
                CurrentEmployee = i.Item1;
                CurrentEmployeeName = i.Item2;
            };

            var initialDirectory = string.IsNullOrWhiteSpace(Settings.Default.InitialSaveDirectory) ? AppDomain.CurrentDomain.BaseDirectory : Settings.Default.InitialSaveDirectory;

            var path = _ioDialogs.OpenDirectory("Wybierz folder do zapisu", initialDirectory);

            var image = Resources.vf_logo300x300;

            _progressDialogController = await _dialogCoordinator.ShowProgressAsync(this, "Export raportu do excel", "test");
            _progressDialogController.Maximum = Report.Employees.Count;
            _progressDialogController.SetCancelable(false);

            var success = await Task.Run(() => SaveExcelVertical.SaveWithProgress(Report.Employees.OfType<Employee>(), path, p, image));

            await _progressDialogController.CloseAsync();

            if (!success)
            {
                await _dialogCoordinator.ShowMessageAsync(this, "Operacja exportu", "Wystąpił błąd podczas eksportu");
            }
            else
            {
                Settings.Default.InitialSaveDirectory = path;
                Settings.Default.Save();
            }
        }
    }
}