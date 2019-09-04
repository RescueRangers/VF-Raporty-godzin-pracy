using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Caliburn.Micro;
using DAL;

namespace CM.Reports.ViewModels
{
    class EmployeeDetailsViewModel : PropertyChangedBase
    {
        private ObservableCollection<DayViewModel> _days;
        private string _name;
        private decimal? _totalOvertime2;
        private decimal? _totalOvertime1;
        private decimal? _totalNormalWork;
        private decimal? _totalWorkHours;

        public ObservableCollection<DayViewModel> Days
        {
            get => _days;
            set
            {
                if (Equals(value, _days)) return;
                _days = value;
                NotifyOfPropertyChange(() => Days);
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                if (value == _name) return;
                _name = value;
                NotifyOfPropertyChange(() => Name);
            }
        }

        public decimal? TotalWorkHours
        {
            get => _totalWorkHours;
            set
            {
                if (value == _totalWorkHours) return;
                _totalWorkHours = value;
                NotifyOfPropertyChange(() => TotalWorkHours);
            }
        }

        public decimal? TotalNormalWork
        {
            get => _totalNormalWork;
            set
            {
                if (value == _totalNormalWork) return;
                _totalNormalWork = value;
                NotifyOfPropertyChange(() => TotalNormalWork);
            }
        }

        public decimal? TotalOvertime1
        {
            get => _totalOvertime1;
            set
            {
                if (value == _totalOvertime1) return;
                _totalOvertime1 = value;
                NotifyOfPropertyChange(() => TotalOvertime1);
            }
        }

        public decimal? TotalOvertime2
        {
            get => _totalOvertime2;
            set
            {
                if (value == _totalOvertime2) return;
                _totalOvertime2 = value;
                NotifyOfPropertyChange(() => TotalOvertime2);
            }
        }

        public void MapData(Employee employee)
        {
            var reportDate = employee.Days.First().Date;

            var days = employee.Days.Select(d => d.Date.Day);
            var notWorkingDays =
                Enumerable.Range(1, DateTime.DaysInMonth(reportDate.Year, reportDate.Month)).Except(days);

            var nonWorkingDayViewModels =
                notWorkingDays.Select(d => new DayViewModel(new DateTime(reportDate.Year, reportDate.Month, d)));
            var workingDayViewModels = employee.Days.Select(d => new DayViewModel(d));
            var dayViewModels = nonWorkingDayViewModels.Concat(workingDayViewModels);

            Days = new ObservableCollection<DayViewModel>(dayViewModels.OrderBy(d => d.Date));
            Name = employee.FullName;

            TotalWorkHours = employee.Days.Sum(d => d.WorkHour);
            TotalOvertime1 = employee.Days.Sum(d => d.Overtime50);
            TotalOvertime2 = employee.Days.Sum(d => d.Overtime100);
            TotalNormalWork = employee.Days.Sum(d => d.WorkHour - (d.Overtime50 ?? 0 + d.Overtime100 ?? 0));
        }
    }
}
