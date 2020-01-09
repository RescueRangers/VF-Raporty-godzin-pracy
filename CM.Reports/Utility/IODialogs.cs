using System;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace CM.Reports.Utility
{
    internal class IODialogs : IIODialogs
    {
        /// <summary>
        /// Displays a directory picker dialog
        /// </summary>
        /// <param name="title">Dialog title</param>
        /// <param name="baseDir">Initial directory</param>
        /// <returns>Returns directory path if user picks a directory, otherwise returns null</returns>
        public string OpenDirectory(string title, string baseDir)
        {
            var dialog = new CommonOpenFileDialog { IsFolderPicker = true, InitialDirectory = baseDir, Title = title };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dialog.FileName;
            }

            return null;
        }

        public Task<string> OpenDirectoryAsync(string title, string baseDir)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Displays a standard windows open file dialog with filter set to excel .xls or .xlsx files
        /// </summary>
        /// <param name="title">Dialog title</param>
        /// <param name="baseDir">Initial directory</param>
        /// <returns>Returns File path if user picks a file, otherwise return null</returns>
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

            if (dialog.ShowDialog() == true)
            {
                return dialog.FileName;
            }

            return null;
        }
    }
}