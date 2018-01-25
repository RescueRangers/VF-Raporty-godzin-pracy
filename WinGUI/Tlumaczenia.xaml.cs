using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WinGUI
{
    /// <summary>
    /// Interaction logic for Tlumaczenia.xaml
    /// </summary>
    public partial class Tlumaczenia : Window
    {
        public string Przetlumaczone { get { return doPrzetlumaczenia.Text; } }

        public Tlumaczenia(string oryginal, string tlumaczenie)
        {
            InitializeComponent();
            Oryginal.Content = oryginal;
            doPrzetlumaczenia.Text = tlumaczenie;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
