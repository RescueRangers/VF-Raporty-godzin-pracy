using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using CM.Reports.ViewModels;
using WinGUI_Avalonia.Utility;

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
        #endregion

        #region Overrides

        protected override void Configure()
        {
            _container.Instance<IWindowManager>(new WindowManager());
            _container.Instance<IIODialogs>(new IODialogs());

            _container.PerRequest<ReportViewModel>();
            _container.PerRequest<MainWindowViewModel>();
            _container.PerRequest<EmployeeDetailsViewModel>();
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
        #endregion
    }
}
