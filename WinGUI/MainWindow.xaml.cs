using System.Windows;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            Application.Current.MainWindow = this;
        }
    }
}