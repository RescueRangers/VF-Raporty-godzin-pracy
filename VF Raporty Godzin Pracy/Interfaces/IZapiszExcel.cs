using System.Collections.Generic;
using System.Threading.Tasks;

namespace VF_Raporty_Godzin_Pracy.Interfaces
{
    public interface IZapiszExcel
    {
        /// <summary>
        /// Zapisuje wybranych pracowników do pliku
        /// </summary>
        /// <param name="raport">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu);

        /// <summary>
        /// Zapisuje raporty wszystkich pracowników do oddzielnych plików
        /// </summary>
        /// <param name="raport">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        /// <param name="nazwaPracownika">Lista pracowników do przetworzenia</param>
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu, List<Pracowik> nazwaPracownika);

        /// <summary>
        /// Zapisuje raporty wybranego pracownika do pliku
        /// </summary>
        /// <param name="raport">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        /// <param name="pracownik">Pracownik do raportu</param>
        Task<string> ZapiszDoExcel(Raport raport, string folderDoZapisu, Pracowik pracownik);
    }
}
