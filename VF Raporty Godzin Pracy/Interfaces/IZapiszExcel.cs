using System.Collections.Generic;
using System.Threading.Tasks;

namespace VF_Raporty_Godzin_Pracy.Interfaces
{
    public interface IZapiszExcel
    {
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu);
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu, List<Pracowik> nazwaPracownika);
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu, Pracowik pracownik);
    }
}
