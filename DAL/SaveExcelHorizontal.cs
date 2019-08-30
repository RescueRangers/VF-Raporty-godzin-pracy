using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DAL.Interfaces;
using DAL.Messages;
using OfficeOpenXml;

namespace DAL
{
    /// <summary>
    /// Zapisuje raporty poszczegolnych pracowniku w formacie poziomej tabelki
    /// </summary>
    public class SaveExcelHorizontal : ISaveExcel
    {
        /// <summary>
        /// Zapisuje raporty wszystkich pracowników do oddzielnych plików
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        /// <param name="nazwaPracownika">Lista pracowników do przetworzenia</param>
        public async Task<string> SaveExcel(Report report, string folderDoZapisu, List<Employee> nazwaPracownika)
        {
            var zapiszRaport = Zapisz(report, folderDoZapisu, nazwaPracownika);
            return await zapiszRaport;
        }

        /// <summary>
        /// Zapisuje raporty wybranego pracownika do pliku
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        /// <param name="employee">Pracownik do raportu</param>
        public async Task<string> SaveExcel(Report report, string folderDoZapisu, Employee employee)
        {
            var zapiszRaport = Zapisz(report, folderDoZapisu, new List<Employee>{employee});
            return await zapiszRaport;
        }

        /// <summary>
        /// Zapisuje wybranych pracowników do pliku
        /// </summary>
        /// <param name="report">Raport z ktorego będą zapisywane wyciągi godzin pracowników</param>
        /// <param name="folderDoZapisu">Folder do zapisu raportów</param>
        public async Task<string> SaveExcel(Report report, string folderDoZapisu)
        {
            var zapiszRaport = Zapisz(report, folderDoZapisu, report.Employees);
            return await zapiszRaport;
        }

        private Task<string> Zapisz(Report report, string folderDoZapisu, List<Employee> nazwaPracownika)
        {
            if (report == null)
            {
                return Task.FromResult("Niepoprawny raport.");
            }
            if (nazwaPracownika == null)
            {
               return Task.FromResult("Nie wybrano pracownika z listy");
            }

            var employeeIndex = 0;

            foreach (var pracownik in nazwaPracownika)
            {
                employeeIndex++;
                var employeeMessage = new CurrentEmployeeMessage
                {
                    CurrentEmployeeName = pracownik.EmployeeName(),
                    CurrentEmployeeNumber = employeeIndex,
                    MaxEmployees = nazwaPracownika.Count
                };
                Messenger.Default.Send<CurrentEmployeeMessage>(employeeMessage);
                var template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template.xlsx";
                var nazwaPliku = $@"{folderDoZapisu}\{pracownik.EmployeeName()}.xlsx";
                var znakiDoWyciecia = new[] { ' ', '\n' };

                using (var excel = new ExcelPackage(new FileInfo(template)))
                {
                    var dlugoscRaportu = report.TranslatedHeaders.Count;
                    var wysokoscRaportu = pracownik.GetDays().Count;

                    //Nazwa pracownika w komorce A1, pozniej jest merge tej komorki na cala dlugosc raportu
                    excel.Workbook.Worksheets[1].Name = pracownik.EmployeeName();

                    var arkusz = excel.Workbook.Worksheets[1];
                    arkusz.Cells[1, 1].Value = pracownik.EmployeeName();
                    arkusz.Cells[1, 1, 1, dlugoscRaportu + 1].Merge = true;

                    var naglowekIndeks = 0;
                    var godziny = 0;
                    

                    arkusz.Cells[2, 1].Value = "Data";

                    //Zapelnianie raportu naglowkami
                    foreach (var naglowek in report.TranslatedHeaders)
                    {
                        if (naglowek.Name.ToLower() == "godziny pracy" || naglowek.Name.ToLower() == "normalpln")
                        {
                            godziny = naglowekIndeks +2;
                        }

                        var tekstNaglowka = "";

                        var slowa = naglowek.Name.Split(' ');

                        //Wstawiam nowe linie jezeli slowo ma 5 lub wiecej liter
                        foreach (var slowo in slowa)
                        {
                            var dlugoscTekstu = +slowo.Length + 1;
                            tekstNaglowka += slowo + " ";
                            if (dlugoscTekstu >= 5)
                            {
                                tekstNaglowka += "\n";
                                dlugoscTekstu = 0;
                            }
                        }


                        arkusz.Cells[2, 2 + naglowekIndeks].Value = tekstNaglowka.Trim(znakiDoWyciecia);
                        naglowekIndeks++;
                    }

                    var dzienIndeks = 0;

                    //Wstawianie dni do pierwszej kolumny
                    foreach (var dzien in pracownik.GetDays())
                    {
                        var godzinaIndeks = 0;
                        arkusz.Cells[3 + dzienIndeks, 1].Value = dzien.Date;
                        arkusz.Cells[3 + dzienIndeks, 1].Style.Numberformat.Format = "dd-mm-yyyy";

                        //Wstawianie godzin do raportu
                        foreach (var godzina in dzien.Hours)
                        {
                            arkusz.Cells[3 + dzienIndeks, 2 + godzinaIndeks].Value = godzina;
                            godzinaIndeks++;
                        }
                        dzienIndeks++;
                    }

                    //Wstawianie dodatkowej kolumny w ktorej beda liczone poprawne godziny
                    arkusz.InsertColumn(godziny+1, 1);
                    arkusz.Cells[2, godziny+1].Value = "Godziny\npracy";
                    arkusz.Cells[2, godziny + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    arkusz.Cells[2, dlugoscRaportu + 2].Value = "Razem";
                    arkusz.Cells[2, dlugoscRaportu + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;


                    //Formula liczaca godziny plus podsumowanie
                    for (int i = 3; i < wysokoscRaportu+3; i++)
                    {
                        //Godziny
                        arkusz.Cells[i, godziny + 1].FormulaR1C1 = $"round(R{i}C{godziny}-R{i}C{godziny+2},0)";
                        
                        //Podsumowanie
                        arkusz.Cells[i, dlugoscRaportu + 2].FormulaR1C1 = $"sum(R{i}C2:R{i}C{godziny-1})+sum(R{i}C{godziny+1}:R{i}C{dlugoscRaportu+1})";
                    }

                    arkusz.Cells[wysokoscRaportu + 3, 1].Value = "Podsumowanie";

                    for (int i = 2; i < dlugoscRaportu+3; i++)
                    {
                        arkusz.Cells[wysokoscRaportu + 3, i].FormulaR1C1 = $"sum(R3C{i}:R{wysokoscRaportu+2}C{i})";
                    }

                    arkusz.Cells.AutoFitColumns();
                    arkusz.Cells[2, 1, 2, dlugoscRaportu + 3].Style.WrapText = true;
                    arkusz.Column(godziny).Hidden = true;

                    //Obramowanie komorek
                    arkusz.Cells[2, 1, wysokoscRaportu + 3, dlugoscRaportu + 2].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    arkusz.Cells[2, 1, wysokoscRaportu + 3, dlugoscRaportu + 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    arkusz.Cells[2, 1, wysokoscRaportu + 3, dlugoscRaportu + 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    arkusz.Cells[2, 1, wysokoscRaportu + 3, dlugoscRaportu + 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;


                    excel.SaveAs(new FileInfo(nazwaPliku));
                }
            }
            return Task.FromResult("Operacja wykonana pomyślnie");
        }
    }
}