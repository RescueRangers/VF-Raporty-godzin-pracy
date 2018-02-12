using WinGUI.ViewModel;

namespace WinGUI
{
    public class ViewModelLocator
    {
        private static readonly MainWindowViewModel mainWindowViewModel = new MainWindowViewModel();

        public static MainWindowViewModel MainWindowViewModel => mainWindowViewModel;
    }
}
