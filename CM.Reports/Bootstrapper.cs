using System;
using System.Collections.Generic;
using System.Windows;
using Caliburn.Micro;
using CM.Reports.Utility;
using CM.Reports.ViewModels;
using MahApps.Metro.Controls.Dialogs;

namespace CM.Reports
{
    public class Bootstrapper : BootstrapperBase
    {
        private SimpleContainer _container = new SimpleContainer();

        #region Constructor

        public Bootstrapper()
        {
            Initialize();
        }

        #endregion Constructor

        #region Overrides

        protected override void Configure()
        {
            _container.Instance<IWindowManager>(new WindowManager());
            _container.Instance<IIODialogs>(new IODialogs());
            _container.Instance<IDialogCoordinator>(DialogCoordinator.Instance);

            _container.PerRequest<ReportViewModel>();
            _container.PerRequest<MainWindowViewModel>();
            _container.PerRequest<EmployeeDetailsViewModel>();
            _container.PerRequest<TranslationsViewModel>();
        }

        protected override object GetInstance(Type serviceType, string key)
        {
            return _container.GetInstance(serviceType, key);
        }

        protected override IEnumerable<object> GetAllInstances(Type serviceType)
        {
            return _container.GetAllInstances(serviceType);
        }

        protected override void BuildUp(object instance)
        {
            _container.BuildUp(instance);
        }

        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            DisplayRootViewFor<MainWindowViewModel>();
        }

        #endregion Overrides
    }
}