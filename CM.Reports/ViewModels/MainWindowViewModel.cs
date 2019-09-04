using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Caliburn.Micro;
using DAL;
using WinGUI_Avalonia.Utility;

namespace CM.Reports.ViewModels
{
    class MainWindowViewModel : Conductor<ReportViewModel>.Collection.AllActive
    {
        private IWindowManager _windowManager;
        private IIODialogs _ioDialogs;
        private ReportViewModel _report;
        private bool _isBusy = false;

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

        public MainWindowViewModel(IWindowManager windowManager, IIODialogs ioDialogs, ReportViewModel report)
        {
            _windowManager = windowManager;
            _ioDialogs = ioDialogs;
            _report = report;
            Items.Add(_report);
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
    }
}
