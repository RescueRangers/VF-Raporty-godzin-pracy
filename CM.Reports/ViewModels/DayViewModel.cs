using System;
using Caliburn.Micro;
using DAL;

namespace CM.Reports.ViewModels
{
    internal class DayViewModel : PropertyChangedBase
    {
        private bool _isFreeDay;
        private decimal? _workHours;
        private decimal? _overtime1;
        private decimal? _overtime2;
        private string _absence;
        private decimal? _normalWork;
        private string _absenceForegroundColor = "Black";

        public bool IsFreeDay
        {
            get => _isFreeDay;
            set
            {
                if (value == _isFreeDay) return;
                _isFreeDay = value;
                NotifyOfPropertyChange(() => IsFreeDay);
            }
        }

        public decimal? WorkHours
        {
            get => _workHours;
            set
            {
                if (value == _workHours) return;
                _workHours = value;
                NotifyOfPropertyChange(() => WorkHours);
                NotifyOfPropertyChange(() => NormalWork);
            }
        }

        public decimal? Overtime1
        {
            get => _overtime1;
            set
            {
                if (value == _overtime1) return;
                _overtime1 = value;
                NotifyOfPropertyChange(() => Overtime1);
                NotifyOfPropertyChange(() => NormalWork);
            }
        }

        public decimal? Overtime2
        {
            get => _overtime2;
            set
            {
                if (value == _overtime2) return;
                _overtime2 = value;
                NotifyOfPropertyChange(() => Overtime2);
                NotifyOfPropertyChange(() => NormalWork);
            }
        }

        public string Absence
        {
            get => _absence;
            set
            {
                if (value == _absence) return;
                _absence = value;
                NotifyOfPropertyChange(() => Absence);
                NotifyOfPropertyChange(() => IsAbsence);
            }
        }

        public DateTime Date { get; }

        public decimal? NormalWork
        {
            get => _normalWork;
            set
            {
                if (value == _normalWork) return;
                _normalWork = value;
                NotifyOfPropertyChange(() => NormalWork);
            }
        }

        public string AbsenceForegroundColor
        {
            get => _absenceForegroundColor;
            set
            {
                if (value == _absenceForegroundColor) return;
                _absenceForegroundColor = value;
                NotifyOfPropertyChange(() => AbsenceForegroundColor);
            }
        }

        public bool IsAbsence => !string.IsNullOrWhiteSpace(Absence);

        public DayViewModel(Day day)
        {
            Date = day.Date;
            switch (day.WorkType)
            {
                case WorkType.Normal:
                    var roundedHours = Math.Round(day.WorkHour ?? 0);
                    NormalWork = roundedHours;
                    WorkHours = roundedHours;
                    break;
                case WorkType.Overtime1:
                    NormalWork = Math.Round((day.WorkHour ?? 0) - (day.Overtime50 ?? 0));
                    Overtime1 = day.Overtime50;
                    WorkHours = day.WorkHour;
                    break;
                case WorkType.Overtime2:
                    if (day.WorkHour != day.Overtime100)
                    {
                        NormalWork = day.WorkHour - day.Overtime100;
                    }
                    Overtime2 = day.Overtime100;
                    WorkHours = day.WorkHour;
                    break;
                case WorkType.Absence:
                    if (string.IsNullOrWhiteSpace(day.TranslatedAbsence))
                    {
                        Absence = day.Absence;
                        AbsenceForegroundColor = "Red";
                    }
                    else
                    {
                        Absence = day.TranslatedAbsence;
                        AbsenceForegroundColor = "Black";
                    }
                    break;
                case WorkType.Overtimes:
                    NormalWork = Math.Round(day.WorkHour ?? 0 - (day.Overtime50 ?? 0 + day.Overtime100 ?? 0));
                    Overtime1 = day.Overtime50;
                    Overtime2 = day.Overtime100;
                    WorkHours = day.WorkHour;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public DayViewModel(DateTime date)
        {
            Date = date;
            IsFreeDay = true;
        }

        public override bool Equals(object obj)
        {
            return obj is DayViewModel model &&
                   Date == model.Date;
        }

        protected bool Equals(DayViewModel other)
        {
            return Date.Equals(other.Date);
        }

        public override int GetHashCode()
        {
            return Date.GetHashCode();
        }
    }
}