namespace WinGUI.Utility
{
    public interface IWyborPliku
    {
        string OtworzPlik(string tytul, string filtrWyboru, string katalogPoczatkowy);
        string OtworzFolder(string tytul, string katalogPoczatkowy);
    }
}
