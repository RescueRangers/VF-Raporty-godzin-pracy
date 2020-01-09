namespace CM.Reports.Utility
{
    internal interface IIODialogs
    {
        string OpenFile(string title, string baseDir);

        string OpenDirectory(string title, string baseDir);
    }
}