using System.Collections.Generic;
using System.Threading.Tasks;

namespace DAL.Interfaces
{
    public interface ISaveExcel
    {
        /// <summary>
        /// Zapisuje wybranych pracowników do pliku
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="savePath">Folder do zapisu raportów</param>
        Task<string> SaveExcel(Report report, string savePath);

        /// <summary>
        /// Zapisuje raporty wszystkich pracowników do oddzielnych plików
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="savePath">Folder do zapisu raportów</param>
        /// <param name="employeeName">Lista pracowników do przetworzenia</param>
        Task<string> SaveExcel(string savePath, List<Employee> employeeName);

        /// <summary>
        /// Zapisuje raporty wybranego pracownika do pliku
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="savePath">Folder do zapisu raportów</param>
        /// <param name="employee">Pracownik do raportu</param>
        Task<string> SaveExcel(string savePath, Employee employee);
    }
}