using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using Caliburn.Micro;
using DAL;

namespace CM.Reports.ViewModels
{
    class ReportViewModel : PropertyChangedBase
    {
        private bool _areHeadersTranslated;
        private IWindowManager _windowManager;
        private List<Translation> _notTranslatedHeaders;
        private List<Header> _headers;
        private CollectionView _employees;
        private ObservableCollection<Employee> _employeeCollection;

        public CollectionView Employees
        {
            get => _employees;
            set
            {
                if (Equals(value, _employees)) return;
                _employees = value;
                NotifyOfPropertyChange(() => Employees);
            }
        }

        private bool _isInitialized;
        private Employee _selectedEmployee;
        private string _search;

        public ReportViewModel(IWindowManager windowManager)
        {
            _windowManager = windowManager;
        }

        public List<Header> Headers
        {
            get => _headers;
            set
            {
                if (Equals(value, _headers)) return;
                _headers = value;
                NotifyOfPropertyChange(() => Headers);
            }
        }

        public string Search
        {
            get => _search;
            set
            {
                if (value == _search) return;
                _search = value;
                NotifyOfPropertyChange(() => Search);

                if (!_isInitialized) return;
                FilterCollection(value);
                Employees.Refresh();
            }
        }

        private void FilterCollection(string filter)
        {
            Employees.Filter = o =>
            {
                //Zwraca całą listę jeżeli filtr jest pusty
                if (string.IsNullOrWhiteSpace(filter))
                    return true;

                //Zwraca tych pracowników których imię lub nazwisko zawiera filrowany ciąg znaków
                return o is Employee employee &&
                       employee.FullName.IndexOf(filter, StringComparison.InvariantCultureIgnoreCase) >= 0;
            };
        }

        public Employee SelectedEmployee
        {
            get => _selectedEmployee;
            set
            {
                if (Equals(value, _selectedEmployee)) return;
                _selectedEmployee = value;
                NotifyOfPropertyChange(() => SelectedEmployee);
            }
        }

        public List<Translation> NotTranslatedHeaders
        {
            get => _notTranslatedHeaders;
            set
            {
                if (Equals(value, _notTranslatedHeaders)) return;
                _notTranslatedHeaders = value;
                NotifyOfPropertyChange(() => NotTranslatedHeaders);
                NotifyOfPropertyChange(() => AreHeadersTranslated);
                NotifyOfPropertyChange(() => IconColor);
                NotifyOfPropertyChange(() => Icon);
            }
        }

        public string IconColor
        {
            get
            {
                if (!_isInitialized)
                {
                    return "Gray";
                }
                return AreHeadersTranslated ? "Green" : "Red";
            }
        }

        public string Icon
        {
            get
            {
                if (!_isInitialized)
                {
                    return "ThumbsUpDown";
                }
                return AreHeadersTranslated ? "ThumbUp" : "ThumbDown";
            }
        }

        private bool AreHeadersTranslated => _notTranslatedHeaders.Count == 0;

        public void MapData(Report report)
        {
            _isInitialized = true;
            Headers = report.Headers;
            NotTranslatedHeaders = report.NotTranslatedHeaders;
            Employees = CollectionViewSource.GetDefaultView(report.Employees) as CollectionView;
            Employees?.SortDescriptions.Add(new SortDescription("LastName", ListSortDirection.Ascending));
            Employees?.Refresh();
        }

        

        public void OpenEmployeeDetails()
        {
            if(SelectedEmployee == null) return;

            var details = new EmployeeDetailsViewModel();
            details.MapData(SelectedEmployee);

            _windowManager.ShowWindow(details);
        }
    }
}
