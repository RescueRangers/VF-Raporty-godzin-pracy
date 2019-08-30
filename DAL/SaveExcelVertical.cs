using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DAL.Interfaces;
using DAL.Messages;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DAL
{
    public class SaveExcelVertical : ISaveExcel
    {
        public async Task<string> SaveExcel(Report report, string folderDoZapisu)
        {
            var zapiszRaport = Zapisz(report, report.Employees, folderDoZapisu);
            return await zapiszRaport;
        }

        public async Task<string> SaveExcel(Report report, string folderDoZapisu, List<Employee> nazwaPracownika)
        {
            var zapiszRaport = Zapisz(report, nazwaPracownika, folderDoZapisu);
            return await zapiszRaport;
        }

        public async Task<string> SaveExcel(Report report, string folderDoZapisu, Employee employee)
        {
            var zapiszRaport = Zapisz(report, new List<Employee>{employee}, folderDoZapisu);
            return await zapiszRaport;
        }

        private Task<string> Zapisz(Report report, List<Employee> nazwaPracownika, string folderDoZapisu)
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
                var index = employeeIndex;
                SendMessage(pracownik.EmployeeName(), index, nazwaPracownika.Count);
                

                var template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template_pion.xlsx";
                var nazwaPliku = $@"{folderDoZapisu}\{pracownik.EmployeeName()}.xlsx";
                var znakiDoWyciecia = new[] { ' ', '\n' };

                using (var excel = new ExcelPackage(new FileInfo(template)))
                {
                    //var dlugoscRaportu = raport.TlumaczoneNaglowki.Count;
                    //var wysokoscRaportu = pracownik.GetDni().Count;

                    excel.Workbook.Worksheets[1].Name = pracownik.EmployeeName();

                    var miesiac = pracownik.Days[0].Date.Month;
                    var rok = pracownik.Days[0].Date.Year;

                    var arkusz = excel.Workbook.Worksheets[1];
                    arkusz.Cells[1, 1].Value = "Wykaz godzin pracy - " + pracownik.Days[0].Date.ToString("MMMM", new CultureInfo("pl-PL"));
                    arkusz.Cells[4, 1].Value = pracownik.EmployeeName();

                    var dniPracujace = pracownik.Days.Select(dzien => dzien.Date.Day).ToList();
                    var dniWMiesiacu = Enumerable.Range(1, DateTime.DaysInMonth(rok, miesiac)).ToList();
                    var dniNiePracujace = dniWMiesiacu.Except(dniPracujace).ToList();

                    for (var i = 1; i <= DateTime.DaysInMonth(rok, miesiac); i++)
                    {
                        arkusz.Cells[6 + i, 1].Value = $"{rok}-{miesiac:00}-{i:00}";
                    }

                    var indeksGodzinPracy = report.TranslatedHeaders.IndexOf(report.TranslatedHeaders.Find(naglowek =>
                        naglowek.Name.ToLower() == "normalpln" || naglowek.Name.ToLower() == "godziny pracy"));
                    var indeksNadgodziny100 = report.TranslatedHeaders.IndexOf(report.TranslatedHeaders.Find(
                        naglowek =>
                            string.Equals(naglowek.Name, "NADGODZINY2", StringComparison.InvariantCultureIgnoreCase) ||
                            string.Equals(naglowek.Name, "Nadgodziny 100%", StringComparison.InvariantCultureIgnoreCase)));


                    foreach (var dzien in pracownik.Days)
                    {
                        var numerDnia = dzien.Date.Day;
                        var indeksyGodzin = new List<int>();

                        var godzinyWhere = dzien.Hours.Where(godzina => godzina > 0).ToList();

                        foreach (var godzina in godzinyWhere)
                        {
                            indeksyGodzin.Add(dzien.Hours.IndexOf(godzina));
                        }

                        //Jezeli w dniu wystepuje dwa rodzaje godzin pracy
                        if (indeksyGodzin.Count == 2)
                        {
                            var godzinyPracy = godzinyWhere[1] - godzinyWhere[0];

                            if (indeksyGodzin[1] == indeksNadgodziny100)
                            {
                                godzinyPracy = Convert.ToInt32(godzinyPracy);
                                arkusz.Cells[6 + numerDnia, 2].Value = godzinyPracy;
                                arkusz.Cells[6 + numerDnia, 4].Value = godzinyWhere[0];
                                arkusz.Cells[6 + numerDnia, 5].Value = godzinyPracy + godzinyWhere[0];
                            }
                            else
                            {
                                godzinyPracy = Convert.ToInt32(godzinyPracy);
                                arkusz.Cells[6 + numerDnia, 2].Value = godzinyPracy;
                                arkusz.Cells[6 + numerDnia, 3].Value = godzinyWhere[0];
                                arkusz.Cells[6 + numerDnia, 5].Value = godzinyPracy + godzinyWhere[0];
                            }
                        }
                        //Jezeli w dniu wystepuje godziny 50%, 100% i normalne godziny
                        else if (indeksyGodzin.Count == 3)
                        {
                            var godzinyPracy = godzinyWhere[0] - godzinyWhere[1];
                            godzinyPracy = Convert.ToInt32(godzinyPracy);
                            arkusz.Cells[6 + numerDnia, 2].Value = godzinyPracy;
                            arkusz.Cells[6 + numerDnia, 3].Value = godzinyWhere[1];
                            arkusz.Cells[6 + numerDnia, 4].Value = godzinyWhere[2];
                            arkusz.Cells[6 + numerDnia, 5].Value = godzinyPracy + godzinyWhere[1] + godzinyWhere[2];

                        }
                        //Tylko jeden typ godzin, wypelnia tabelke albo godzinami albo nazwa naglowka
                        else
                        {
                            //Godziny pracy
                            if (indeksyGodzin[0] == indeksGodzinPracy)
                            {

                                var godzinyPracy = Convert.ToInt32(godzinyWhere[0]);

                                arkusz.Cells[6 + numerDnia, 2].Value = godzinyPracy;
                                arkusz.Cells[6 + numerDnia, 5].Value = godzinyPracy;
                            }
                            //Nadgodziny 100%
                            else if (indeksyGodzin[0] == indeksNadgodziny100)
                            {
                                arkusz.Cells[6 + numerDnia, 4].Value = godzinyWhere[0];
                                arkusz.Cells[6 + numerDnia, 5].Value = godzinyWhere[0];
                            }
                            //Reszta, nazwy naglowkow
                            else
                            {
                                arkusz.Cells[6 + numerDnia, 2].Value = report.TranslatedHeaders[indeksyGodzin[0]].Name;
                                arkusz.Cells[6 + numerDnia, 2, 6 + numerDnia, 5].Merge = true;
                            }
                        }

                        foreach (var dzienNiePracujacy in dniNiePracujace)
                        {
                            arkusz.Cells[6 + dzienNiePracujacy, 2, 6 + dzienNiePracujacy, 5].Merge = true;
                            arkusz.Cells[6 + dzienNiePracujacy, 2, 6 + dzienNiePracujacy, 5].Style.Border.Diagonal.Style =
                                ExcelBorderStyle.Thin;
                            arkusz.Cells[6 + dzienNiePracujacy, 2, 6 + dzienNiePracujacy, 5].Style.Border.DiagonalDown =
                                true;
                        }

                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 1].Value = "Razem";
                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 2].FormulaR1C1 = $"sum(R7C2:R{dniWMiesiacu.Count + 6}C2)";

                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 3].FormulaR1C1 = $"sum(R7C3:R{dniWMiesiacu.Count + 6}C3)";
                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 4].FormulaR1C1 = $"sum(R7C4:R{dniWMiesiacu.Count + 6}C4)";
                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 5].FormulaR1C1 = $"sum(R7C5:R{dniWMiesiacu.Count + 6}C5)";

                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 2].Style.Numberformat.Format = "#";
                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 3, 6 + dniWMiesiacu.Count + 1, 5].Style.Numberformat
                            .Format = "0.00";
                        arkusz.Cells[6 + dniWMiesiacu.Count + 1, 2, 6 + dniWMiesiacu.Count + 1, 5].Style.HorizontalAlignment =
                            ExcelHorizontalAlignment.Center;

                        arkusz.Cells[6, 1, 7 + dniWMiesiacu.Count, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        arkusz.Cells[6, 1, 7 + dniWMiesiacu.Count, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        arkusz.Cells[6, 1, 7 + dniWMiesiacu.Count, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        arkusz.Cells[6, 1, 7 + dniWMiesiacu.Count, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    excel.SaveAs(new FileInfo(nazwaPliku));
                }
            }
            return Task.FromResult("Operacja wykonana pomyślnie");
        }

        private void SendMessage(string nazwaPracownika, int employeeIndex, int maxEmployees)
        {
            Messenger.Default.Send<CurrentEmployeeMessage>(new CurrentEmployeeMessage
            {
                CurrentEmployeeName = nazwaPracownika,
                CurrentEmployeeNumber = employeeIndex,
                MaxEmployees = maxEmployees
            });
        }
    }
}
