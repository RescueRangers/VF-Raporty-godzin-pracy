using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinGUI_Avalonia.Utility
{
    interface IIODialogs
    {
        string OpenFile(string title, string baseDir);
        string OpenDirectory(string title, string baseDir);
    }
}
