using System.Windows;
using System.Windows.Controls;

namespace CM.Reports.Views
{
    /// <summary>
    /// Interaction logic for ReportView.xaml
    /// </summary>
    public partial class ReportView : UserControl
    {
        public ReportView()
        {
            InitializeComponent();
        }

        private void Employees_OnDragEnter(object sender, DragEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                dropCanvas.Visibility = Visibility.Visible;
                //grid.BorderBrush = Brushes.Blue;
                //grid.BorderThickness = new Thickness(2);
            }
        }

        private void Employees_OnDragLeave(object sender, DragEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                dropCanvas.Visibility = Visibility.Hidden;
                //grid.BorderBrush = Brushes.Transparent;
                //grid.BorderThickness = new Thickness(0);
            }
        }

        private void Employees_OnDrop(object sender, DragEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                dropCanvas.Visibility = Visibility.Hidden;
                //grid.BorderBrush = Brushes.Transparent;
                //grid.BorderThickness = new Thickness(0);
            }
        }
    }
}
