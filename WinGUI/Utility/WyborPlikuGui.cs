using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace WinGUI.Utility
{
    public class WyborPlikuGui : IWyborPliku
    {
        public string OtworzPlik(string tytul, string filtrWyboru, string katalogPoczatkowy)
        {
            var wybranyPlik = string.Empty;
            var otworzPlik = new OpenFileDialog
            {
                Title = tytul,
                Filter = filtrWyboru,
                InitialDirectory = katalogPoczatkowy
            };
            if (otworzPlik.ShowDialog() == true)
            {
                wybranyPlik = otworzPlik.FileName;
            }

            return string.IsNullOrWhiteSpace(wybranyPlik) ? null : wybranyPlik;
        }

        public string OtworzFolder(string tytul, string katalogPoczatkowy)
        {
            string folderDoZapisu;
            var wyborFolderu = new CommonOpenFileDialog
            {
                Title = tytul,
                IsFolderPicker = true,
                InitialDirectory = katalogPoczatkowy
            };

            if (wyborFolderu.ShowDialog() == CommonFileDialogResult.Ok)
            {
                folderDoZapisu = wyborFolderu.FileName;
            }
            else
            {
                return null;
            }

            return folderDoZapisu;
        }
    }
}