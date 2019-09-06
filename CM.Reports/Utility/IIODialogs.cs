namespace CM.Reports.Utility
{
    interface IIODialogs
    {
        string OpenFile(string title, string baseDir);
        string OpenDirectory(string title, string baseDir);
    }
}
