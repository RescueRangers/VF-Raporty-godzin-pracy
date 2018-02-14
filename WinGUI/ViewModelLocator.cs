using WinGUI.ViewModel;

namespace WinGUI
{
    public class ViewModelLocator
    {
        public static MainWindowViewModel MainWindowViewModel { get; } = new MainWindowViewModel();
    }
}
