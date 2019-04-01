using System.Windows;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for ProgressDialog.xaml
    /// </summary>
    public partial class ProgressDialog : Window
    {
        public ProgressDialog()
        {
            InitializeComponent();
            WindowStyle = WindowStyle.None;
            Owner = Application.Current.MainWindow;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }
    }
}
