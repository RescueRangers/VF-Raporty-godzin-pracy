using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private readonly string _template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template_pion.xlsx";

        public async Task<string> SaveExcel(Report report, string savePath)
        {
            var saveReport = Save(report, report.Employees, savePath);
            return await saveReport;
        }

        public async Task<string> SaveExcel(Report report, string savePath, List<Employee> employeeName)
        {
            var saveReport = Save(report, employeeName, savePath);
            return await saveReport;
        }

        public async Task<string> SaveExcel(Report report, string savePath, Employee employee)
        {
            var saveReport = Save(report, new List<Employee>{employee}, savePath);
            return await saveReport;
        }

        private Task<string> Save(Report report, List<Employee> employees, string savePath)
        {
            if (report == null)
            {
                return Task.FromResult("Niepoprawny raport.");
            }
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

                using (var excel = new ExcelPackage(new FileInfo(_template)))
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

        private static void StyleWorksheet(ExcelWorksheet worksheet, int lastRow)
        {
            worksheet.Cells[lastRow + 1, 2].Style.Numberformat.Format = "0.00";
            worksheet.Cells[lastRow + 1, 3, lastRow + 1, 5].Style.Numberformat
                .Format = "0.00";
            worksheet.Cells[lastRow + 1, 2, lastRow + 1, 5].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;

            worksheet.Cells[6, 1, 1 + lastRow, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[6, 1, 1 + lastRow, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[6, 1, 1 + lastRow, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[6, 1, 1 + lastRow, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
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
                    worksheet.Cells[row, 4].Value = day.Overtime100;
                    worksheet.Cells[row, 5].Value = day.Overtime100;
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
