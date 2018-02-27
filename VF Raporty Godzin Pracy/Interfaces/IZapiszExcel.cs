using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy.Interfaces
{
    public interface IZapiszExcel
    {
        string ZapiszDoExcel(Raport raport, string folderDoZapisu);
        string ZapiszDoExcel(Raport raport, string folderDoZapisu, List<Pracowik> nazwaPracownika);
    }
}
