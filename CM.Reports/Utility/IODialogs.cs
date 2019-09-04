using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace WinGUI_Avalonia.Utility
{
    class IODialogs : IIODialogs
    {
        public string OpenDirectory(string title, string baseDir)
        {
            throw new NotImplementedException();
        }

        public Task<string> OpenDirectoryAsync(string title, string baseDir)
        {
            throw new NotImplementedException();
        }

        public string OpenFile(string title, string baseDir)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = false,
                AddExtension = false,
                Filter = "Pliki Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                InitialDirectory = baseDir,
                Title = title
            };

            if (dialog.ShowDialog() == true )
            {
                return dialog.FileName;
            }

            return null;
        }

        //public async Task<string> OpenFileAsync(string title, List<FileDialogFilter> filters, string baseDir)
        //{
        //    var dialog = new OpenFileDialog {AllowMultiple = false, InitialDirectory = baseDir, Title = title, Filters = filters};
        //    var @return = await dialog.ShowAsync(Application.Current.MainWindow);

        //    return @return.FirstOrDefault();
        //}
    }
}
