using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DAL.Extensions;
using DAL.Interfaces;
using DAL.Messages;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DAL
{
    public class SaveExcelVertical : ISaveExcel
    {
        private static readonly string Template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template_pion.xlsx";

        public async Task<string> SaveExcel(Report report, string savePath)
        {
            var saveReport = Save(report.Employees, savePath);
            return await saveReport;
        }

        public async Task<string> SaveExcel(string savePath, List<Employee> employeeName)
        {
            var saveReport = Save(employeeName, savePath);
            return await saveReport;
        }

        public async Task<string> SaveExcel(string savePath, Employee employee)
        {
            var saveReport = Save(new List<Employee> { employee }, savePath);
            return await saveReport;
        }

        private Task<string> Save(List<Employee> employees, string savePath)
        {
            if (employees == null)
            {
                return Task.FromResult("Nie wybrano pracownika z listy");
            }

            var employeeIndex = 0;

            foreach (var employee in employees)
            {
                employeeIndex++;
                SendMessage(employee.FullName, employeeIndex, employees.Count);

                var fileName = $@"{savePath}\{employee.FullName}.xlsx";

                using (var excel = new ExcelPackage(new FileInfo(Template)))
                {
                    excel.Workbook.Worksheets[1].Name = employee.FullName;

                    var month = employee.Days[0].Date.Month;
                    var year = employee.Days[0].Date.Year;

                    var worksheet = excel.Workbook.Worksheets[1];
                    worksheet.Cells[1, 1].Value = "Wykaz godzin pracy - " + employee.Days[0].Date.ToString("MMMM", new CultureInfo("pl-PL"));
                    worksheet.Cells[4, 1].Value = employee.FullName;

                    var workDays = employee.Days.Select(day => day.Date.Day).ToList();
                    var daysInMonth = Enumerable.Range(1, DateTime.DaysInMonth(year, month)).ToList();
                    var nonWorkDays = daysInMonth.Except(workDays).ToList();

                    for (var i = 1; i <= DateTime.DaysInMonth(year, month); i++)
                    {
                        worksheet.Cells[6 + i, 1].Value = $"{year}-{month:00}-{i:00}";
                    }

                    foreach (var day in employee.Days)
                    {
                        FillWorkHours(worksheet, day);

                        foreach (var nonWorkDay in nonWorkDays)
                        {
                            var row = 6 + nonWorkDay;
                            worksheet.Cells[row, 2, row, 5].Merge = true;
                            worksheet.Cells[row, 2, row, 5].Style.Border.Diagonal.Style =
                                ExcelBorderStyle.Thin;
                            worksheet.Cells[row, 2, row, 5].Style.Border.DiagonalDown =
                                true;
                        }
                    }

                    var lastRow = daysInMonth.Count + 6;
                    worksheet.Cells[lastRow + 1, 1].Value = "Razem";
                    worksheet.Cells[lastRow + 1, 2].FormulaR1C1 = $"sum(R7C2:R{lastRow}C2)";
                    worksheet.Cells[lastRow + 1, 3].FormulaR1C1 = $"sum(R7C3:R{lastRow}C3)";
                    worksheet.Cells[lastRow + 1, 4].FormulaR1C1 = $"sum(R7C4:R{lastRow}C4)";
                    worksheet.Cells[lastRow + 1, 5].FormulaR1C1 = $"sum(R7C5:R{lastRow}C5)";

                    StyleWorksheet(worksheet, lastRow);

                    excel.SaveAs(new FileInfo(fileName));
                }
            }
            return Task.FromResult("Operacja wykonana pomyślnie");
        }

        public static bool SaveWithProgress(IEnumerable<Employee> employees, string savePath, IProgress<Tuple<int, string>> progress, Image logo = null)
        {
            if (employees == null)
            {
                return false;
            }

            var employeeIndex = 0;

            foreach (var employee in employees)
            {
                employeeIndex++;

                var fileName = $@"{savePath}\{employee.FullName}.xlsx";

                using (var excel = new ExcelPackage())
                {
                    var month = employee.Days[0].Date.Month;
                    var year = employee.Days[0].Date.Year;

                    var worksheet = excel.Workbook.Worksheets.Add(employee.FullName);
                    worksheet.HeaderFooter.FirstHeader.RightAlignedText = "&8&\"Calibri,Regular Bold\" Vest-Fiber sp. z o.o.\r\nul. Piłsudskiego 11\r\n74-400 Dębno\r\nNIP: 597-172-75-69";
                    if (logo != null)
                    {
                        worksheet.HeaderFooter.FirstHeader.InsertPicture(logo,
                            PictureAlignment.Left);
                    }
                    worksheet.Cells[1, 1].Value = "Wykaz godzin pracy - " + employee.Days[0].Date.ToString("MMMM", new CultureInfo("pl-PL"));

                    FormatTemplate(worksheet);

                    worksheet.Cells[4, 1].Value = employee.FullName;

                    var workDays = employee.Days.Select(day => day.Date.Day).ToList();
                    var daysInMonth = Enumerable.Range(1, DateTime.DaysInMonth(year, month)).ToList();
                    var nonWorkDays = daysInMonth.Except(workDays).ToList();

                    for (var i = 1; i <= DateTime.DaysInMonth(year, month); i++)
                    {
                        worksheet.Cells[6 + i, 1].Value = $"{year}-{month:00}-{i:00}";
                    }

                    foreach (var day in employee.Days)
                    {
                        FillWorkHours(worksheet, day);

                        foreach (var nonWorkDay in nonWorkDays)
                        {
                            var row = 6 + nonWorkDay;
                            worksheet.Cells[row, 2, row, 5].Merge = true;
                            worksheet.Cells[row, 2, row, 5].Style.Border.Diagonal.Style =
                                ExcelBorderStyle.Thin;
                            worksheet.Cells[row, 2, row, 5].Style.Border.DiagonalDown =
                                true;
                        }
                    }

                    var lastRow = daysInMonth.Count + 6;
                    worksheet.Cells[lastRow + 1, 1].Value = "Razem";
                    worksheet.Cells[lastRow + 1, 2].FormulaR1C1 = $"sum(R7C2:R{lastRow}C2)";
                    worksheet.Cells[lastRow + 1, 3].FormulaR1C1 = $"sum(R7C3:R{lastRow}C3)";
                    worksheet.Cells[lastRow + 1, 4].FormulaR1C1 = $"sum(R7C4:R{lastRow}C4)";
                    worksheet.Cells[lastRow + 1, 5].FormulaR1C1 = $"sum(R7C5:R{lastRow}C5)";

                    StyleWorksheet(worksheet, lastRow);

                    excel.SaveAs(new FileInfo(fileName));

                    progress.Report(new Tuple<int, string>(employeeIndex, employee.FullName));
                }
            }
            return true;
        }

        private static void FormatTemplate(ExcelWorksheet worksheet)
        {
            //Margins
            worksheet.PrinterSettings.TopMargin = 2.5m / 2.54m;
            worksheet.PrinterSettings.LeftMargin = 1.8m / 2.54m;
            worksheet.PrinterSettings.RightMargin = 1.8m / 2.54m;
            worksheet.PrinterSettings.BottomMargin = 1m / 2.54m;
            worksheet.PrinterSettings.HeaderMargin = 0.8m / 2.54m;
            worksheet.PrinterSettings.FooterMargin = 0m;

            //Fill template with text
            worksheet.Cells[2, 1].Value = "Karta pracy";
            worksheet.Cells[6, 1].Value = "Data";
            worksheet.Cells[6, 2].Value = "Normatywny czas\r\npracy w godzinach";
            worksheet.Cells[6, 3].Value = "Godziny 50%";
            worksheet.Cells[6, 4].Value = "Godziny 100%";
            worksheet.Cells[6, 5].Value = "Łączny czas przepracowany \r\nw godzinach";

            //Merge rows
            worksheet.Cells[1, 1, 1, 5].Merge = true;
            worksheet.Cells[2, 1, 2, 5].Merge = true;
            worksheet.Cells[4, 1, 4, 5].Merge = true;

            //Font setup
            worksheet.Cells[1, 1].Style.Font.Name = "Calibri";
            worksheet.Cells[1, 1].Style.Font.Size = 16;
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[2, 1].Style.Font.Name = "Calibri";
            worksheet.Cells[2, 1].Style.Font.Size = 16;
            worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[4, 1].Style.Font.Name = "Calibri";
            worksheet.Cells[4, 1].Style.Font.Size = 16;
            worksheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[6, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[6, 1].Style.WrapText = true;

            worksheet.Cells[6, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[6, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[6, 2].Style.WrapText = true;

            worksheet.Cells[6, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[6, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[6, 3].Style.WrapText = true;

            worksheet.Cells[6, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[6, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[6, 4].Style.WrapText = true;

            worksheet.Cells[6, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[6, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[6, 5].Style.WrapText = true;

            //Row setup
            worksheet.Row(1).CustomHeight = true;
            worksheet.Row(1).Height = 21;

            worksheet.Row(2).CustomHeight = true;
            worksheet.Row(2).Height = 21;

            worksheet.Row(4).CustomHeight = true;
            worksheet.Row(4).Height = 21;

            worksheet.Row(6).CustomHeight = true;
            worksheet.Row(6).Height = 31.5;

            //Column setup
            worksheet.Column(1).SetTrueColumnWidth(10.86);
            worksheet.Column(2).SetTrueColumnWidth(17.57);
            worksheet.Column(3).SetTrueColumnWidth(12);
            worksheet.Column(4).SetTrueColumnWidth(13.29);
            worksheet.Column(5).SetTrueColumnWidth(28.71);

            //Employee signature
            worksheet.Cells[44, 4, 44, 5].Merge = true;
            worksheet.Cells[44, 4, 44, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;

            worksheet.Cells[45, 4, 45, 5].Merge = true;
            worksheet.Cells[45, 4, 45, 5].Value = "Podpis pracownika";
            worksheet.Cells[45, 4, 45, 5].Style.Font.Size = 11;
            worksheet.Cells[45, 4, 45, 5].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
            worksheet.Cells[45, 4, 45, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private static void StyleWorksheet(ExcelWorksheet worksheet, int lastRow)
        {
            worksheet.Cells[7, 2, lastRow, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells[lastRow + 1, 2].Style.Numberformat.Format = "0.00";
            worksheet.Cells[lastRow + 1, 3, lastRow + 1, 5].Style.Numberformat
                .Format = "0.00";
            worksheet.Cells[lastRow + 1, 2, lastRow + 1, 5].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;

            var range = worksheet.Cells[6, 1, 1 + lastRow, 5];

            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Font.Size = 12;
        }

        private static void FillWorkHours(ExcelWorksheet worksheet, Day day)
        {
            var row = day.Date.Day + 6;

            switch (day.WorkType)
            {
                case WorkType.Normal:
                    if (day.WorkHour != null)
                    {
                        worksheet.Cells[row, 2].Value = Math.Round(day.WorkHour.Value);
                        worksheet.Cells[row, 5].Value = Math.Round(day.WorkHour.Value);
                    }

                    break;

                case WorkType.Overtime1:
                    if (day.WorkHour != null && day.Overtime50 != null)
                    {
                        worksheet.Cells[row, 2].Value = Math.Round(day.WorkHour.Value - day.Overtime50.Value);
                        worksheet.Cells[row, 3].Value = day.Overtime50;
                        worksheet.Cells[row, 5].Value = day.WorkHour;
                    }

                    break;

                case WorkType.Overtime2:
                    if (day.WorkHour != day.Overtime100)
                    {
                        worksheet.Cells[row, 2].Value = Math.Round(day.WorkHour.Value - day.Overtime100.Value);
                    }
                    else
                    {
                        worksheet.Cells[row, 2].Value = 0;
                    }
                    worksheet.Cells[row, 4].Value = day.Overtime100;
                    worksheet.Cells[row, 5].Value = day.WorkHour;
                    break;

                case WorkType.Absence:
                    worksheet.Cells[row, 2].Value = day.TranslatedAbsence;
                    worksheet.Cells[row, 2, row, 5].Merge = true;
                    break;

                case WorkType.Overtimes:

                    if (day.WorkHour.HasValue && day.Overtime50.HasValue && day.Overtime100.HasValue)
                    {
                        worksheet.Cells[row, 2].Value =
                            Math.Round(day.WorkHour.Value - day.Overtime50.Value - day.Overtime100.Value);
                        worksheet.Cells[row, 3].Value = day.Overtime50;
                        worksheet.Cells[row, 4].Value = day.Overtime100;
                        worksheet.Cells[row, 5].Value = day.WorkHour;
                    }

                    break;
            }
        }

        private void SendMessage(string employeeName, int employeeIndex, int maxEmployees)
        {
            Messenger.Default.Send(new CurrentEmployeeMessage
            {
                CurrentEmployeeName = employeeName,
                CurrentEmployeeNumber = employeeIndex,
                MaxEmployees = maxEmployees
            });
        }
    }
}