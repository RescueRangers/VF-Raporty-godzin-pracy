using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Caliburn.Micro;
using DAL;
using MahApps.Metro.Controls.Dialogs;
using WinGUI_Avalonia.Utility;

namespace CM.Reports.ViewModels
{
    class MainWindowViewModel : Conductor<ReportViewModel>.Collection.AllActive
    {
        private IWindowManager _windowManager;
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

        public MainWindowViewModel(IWindowManager windowManager, IIODialogs ioDialogs, ReportViewModel report, IDialogCoordinator dialogCoordinator)
        {
            _windowManager = windowManager;
            _ioDialogs = ioDialogs;
            _report = report;
            Items.Add(_report);
            _dialogCoordinator = dialogCoordinator;
        }

        public  async Task OpenExcelReport()
        {
            var filePath = _ioDialogs.OpenFile("Otwórz plik z eksportem godzin pracy.",
                Environment.SpecialFolder.MyDocuments.ToString());

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                IsBusy = true;
                var report = await Task.Run(() => DAL.Report.Create(filePath));
                IsBusy = false;
                _report.MapData(report);
            }
        }

        public void OpenTranslations()
        {
            var translations = new TranslationsViewModel();

            if(Report.NotTranslatedHeaders != null) translations.HeadersToTranslate = new ObservableCollection<Translation>(Report.NotTranslatedHeaders);

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

            var image = Properties.Resources.vf_logo300x300;

            _progressDialogController = await _dialogCoordinator.ShowProgressAsync(this, "Export raportu do excel", "test");
            _progressDialogController.Maximum = Report.Employees.Count;
            _progressDialogController.SetCancelable(false);

            var success = await Task.Run(() => SaveExcelVertical.SaveWithProgress(Report.Employees.OfType<Employee>(),@"C:\Users\user\Desktop\Roczna KARTA PRACY 2017\Test",p, image));

            await _progressDialogController.CloseAsync();

            if (!success)
            {
                await _dialogCoordinator.ShowMessageAsync(this, "Operacja exportu", "Wystąpił błąd podczas eksportu");
            }
        }
    }
}
