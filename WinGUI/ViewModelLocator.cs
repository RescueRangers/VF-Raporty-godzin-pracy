using System;
using WinGUI.Servicess;
using WinGUI.ViewModel;

namespace WinGUI
{
    public class ViewModelLocator
    {
        private static readonly IProgressDialogService ProgressDialogService = new ProgressDialogService();

        public static MainWindowViewModel MainWindowViewModel
        {
            get { return (MainWindowViewModel) Activator.CreateInstance(typeof(MainWindowViewModel), ProgressDialogService); }
        }

        
    }
}
