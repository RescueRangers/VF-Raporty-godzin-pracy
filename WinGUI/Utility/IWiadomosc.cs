namespace WinGUI.Utility
{
    public interface IWiadomosc
    {
        void WyslijWiadomosc(string tresc, string naglowek, TypyWiadomosci typWiadomosci);

    }

    public enum TypyWiadomosci
    {
        Informacja,
        Blad
    }
}
