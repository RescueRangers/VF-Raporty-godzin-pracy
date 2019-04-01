using System.Windows;

namespace WinGUI.Utility
{
    public class WiadomoscGui : IWiadomosc
    {
        public void WyslijWiadomosc(string tresc, string naglowek, TypyWiadomosci typWiadomosci)
        {
            if (typWiadomosci == TypyWiadomosci.Informacja)
            {
                MessageBox.Show(tresc, naglowek, MessageBoxButton.OK, MessageBoxImage.Information);
            }

            if (typWiadomosci == TypyWiadomosci.Blad)
            {
                MessageBox.Show(tresc, naglowek, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}