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
                var index = employeeIndex;
                SendMessage(employee.EmployeeName(), index, employees.Count);

                var template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template_pion.xlsx";
                var fileName = $@"{savePath}\{employee.EmployeeName()}.xlsx";

                using (var excel = new ExcelPackage(new FileInfo(template)))
                {
                    excel.Workbook.Worksheets[1].Name = employee.EmployeeName();

                    var month = employee.Days[0].Date.Month;
                    var year = employee.Days[0].Date.Year;

                    var worksheet = excel.Workbook.Worksheets[1];
                    worksheet.Cells[1, 1].Value = "Wykaz godzin pracy - " + employee.Days[0].Date.ToString("MMMM", new CultureInfo("pl-PL"));
                    worksheet.Cells[4, 1].Value = employee.EmployeeName();

                    var workDays = employee.Days.Select(day => day.Date.Day).ToList();
                    var daysInMonth = Enumerable.Range(1, DateTime.DaysInMonth(year, month)).ToList();
                    var nonWorkDays = daysInMonth.Except(workDays).ToList();

                    for (var i = 1; i <= DateTime.DaysInMonth(year, month); i++)
                    {
                        worksheet.Cells[6 + i, 1].Value = $"{year}-{month:00}-{i:00}";
                    }

                    var workHoursIndex = report.TranslatedHeaders.IndexOf(report.TranslatedHeaders.Find(header =>
                        string.Equals(header.Name, "normalpln", StringComparison.OrdinalIgnoreCase) || string.Equals(header.Name, "godziny pracy", StringComparison.OrdinalIgnoreCase)));
                    var overTime100Index = report.TranslatedHeaders.IndexOf(report.TranslatedHeaders.Find(
                        header =>
                            string.Equals(header.Name, "NADGODZINY2", StringComparison.InvariantCultureIgnoreCase) ||
                            string.Equals(header.Name, "Nadgodziny 100%", StringComparison.InvariantCultureIgnoreCase)));


                    foreach (var day in employee.Days)
                    {
                        var dayNumber = day.Date.Day;
                        //var hoursIndex = new List<int>();

                        //var hours = day.Hours.Where(hour => hour > 0).ToList();

                        //foreach (var hour in hours)
                        //{
                        //    hoursIndex.Add(day.Hours.IndexOf(hour));
                        //}

                        switch (day.WorkType)
                        {
                            case WorkType.Normal:
                                worksheet.Cells[6 + dayNumber, 2].Value = day.WorkHour;
                                break;
                            case WorkType.Overtime1:
                                worksheet.Cells[6 + dayNumber, 2].Value = day.WorkHour - day.Overtime50;
                                worksheet.Cells[6 + dayNumber, 3].Value = day.Overtime50;
                                worksheet.Cells[6 + dayNumber, 5].Value = day.WorkHour;
                                break;
                            case WorkType.Overtime2:
                                worksheet.Cells[6 + dayNumber, 4].Value = day.Overtime100;
                                worksheet.Cells[6 + dayNumber, 5].Value = day.Overtime100;
                                break;
                            case WorkType.Absence:
                                worksheet.Cells[6 + dayNumber, 2].Value = day.Absence;
                                worksheet.Cells[6 + dayNumber, 2, 6 + dayNumber, 5].Merge = true;
                                break;
                            case WorkType.Overtimes:
                                worksheet.Cells[6 + dayNumber, 2].Value = day.WorkHour - day.Overtime50;
                                worksheet.Cells[6 + dayNumber, 3].Value = day.Overtime50;
                                worksheet.Cells[6 + dayNumber, 4].Value = day.Overtime100;
                                worksheet.Cells[6 + dayNumber, 5].Value = day.WorkHour + day.Overtime100;
                                break;
                        }

                        ////Jezeli w dniu wystepuje dwa rodzaje godzin pracy
                        //if (hoursIndex.Count == 2)
                        //{
                        //    var workHours = hours[1] - hours[0];

                        //    if (hoursIndex[1] == overTime100Index)
                        //    {
                        //        workHours = Convert.ToInt32(workHours);
                        //        worksheet.Cells[6 + dayNumber, 2].Value = workHours;
                        //        worksheet.Cells[6 + dayNumber, 4].Value = hours[0];
                        //        worksheet.Cells[6 + dayNumber, 5].Value = workHours + hours[0];
                        //    }
                        //    else
                        //    {
                        //        workHours = Convert.ToInt32(workHours);
                        //        worksheet.Cells[6 + dayNumber, 2].Value = workHours;
                        //        worksheet.Cells[6 + dayNumber, 3].Value = hours[0];
                        //        worksheet.Cells[6 + dayNumber, 5].Value = workHours + hours[0];
                        //    }
                        //}
                        ////Jezeli w dniu wystepuje godziny 50%, 100% i normalne godziny
                        //else if (hoursIndex.Count == 3)
                        //{
                        //    var workHours = hours[0] - hours[1];
                        //    workHours = Convert.ToInt32(workHours);
                        //    worksheet.Cells[6 + dayNumber, 2].Value = workHours;
                        //    worksheet.Cells[6 + dayNumber, 3].Value = hours[1];
                        //    worksheet.Cells[6 + dayNumber, 4].Value = hours[2];
                        //    worksheet.Cells[6 + dayNumber, 5].Value = workHours + hours[1] + hours[2];

                        //}
                        ////Tylko jeden typ godzin, wypelnia tabelke albo godzinami albo nazwa naglowka
                        //else
                        //{
                        //    //Godziny pracy
                        //    if (hoursIndex[0] == workHoursIndex)
                        //    {

                        //        var workhours = Convert.ToInt32(hours[0]);

                        //        worksheet.Cells[6 + dayNumber, 2].Value = workhours;
                        //        worksheet.Cells[6 + dayNumber, 5].Value = workhours;
                        //    }
                        //    //Nadgodziny 100%
                        //    else if (hoursIndex[0] == overTime100Index)
                        //    {
                        //        worksheet.Cells[6 + dayNumber, 4].Value = hours[0];
                        //        worksheet.Cells[6 + dayNumber, 5].Value = hours[0];
                        //    }
                        //    //Reszta, nazwy naglowkow
                        //    else
                        //    {
                        //        worksheet.Cells[6 + dayNumber, 2].Value = report.TranslatedHeaders[hoursIndex[0]].Name;
                        //        worksheet.Cells[6 + dayNumber, 2, 6 + dayNumber, 5].Merge = true;
                        //    }
                        //}

                        foreach (var nonWorkDay in nonWorkDays)
                        {
                            worksheet.Cells[6 + nonWorkDay, 2, 6 + nonWorkDay, 5].Merge = true;
                            worksheet.Cells[6 + nonWorkDay, 2, 6 + nonWorkDay, 5].Style.Border.Diagonal.Style =
                                ExcelBorderStyle.Thin;
                            worksheet.Cells[6 + nonWorkDay, 2, 6 + nonWorkDay, 5].Style.Border.DiagonalDown =
                                true;
                        }

                        worksheet.Cells[6 + daysInMonth.Count + 1, 1].Value = "Razem";
                        worksheet.Cells[6 + daysInMonth.Count + 1, 2].FormulaR1C1 = $"sum(R7C2:R{daysInMonth.Count + 6}C2)";

                        worksheet.Cells[6 + daysInMonth.Count + 1, 3].FormulaR1C1 = $"sum(R7C3:R{daysInMonth.Count + 6}C3)";
                        worksheet.Cells[6 + daysInMonth.Count + 1, 4].FormulaR1C1 = $"sum(R7C4:R{daysInMonth.Count + 6}C4)";
                        worksheet.Cells[6 + daysInMonth.Count + 1, 5].FormulaR1C1 = $"sum(R7C5:R{daysInMonth.Count + 6}C5)";

                        worksheet.Cells[6 + daysInMonth.Count + 1, 2].Style.Numberformat.Format = "#";
                        worksheet.Cells[6 + daysInMonth.Count + 1, 3, 6 + daysInMonth.Count + 1, 5].Style.Numberformat
                            .Format = "0.00";
                        worksheet.Cells[6 + daysInMonth.Count + 1, 2, 6 + daysInMonth.Count + 1, 5].Style.HorizontalAlignment =
                            ExcelHorizontalAlignment.Center;

                        worksheet.Cells[6, 1, 7 + daysInMonth.Count, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[6, 1, 7 + daysInMonth.Count, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[6, 1, 7 + daysInMonth.Count, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[6, 1, 7 + daysInMonth.Count, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    excel.SaveAs(new FileInfo(fileName));
                }
            }
            return Task.FromResult("Operacja wykonana pomyślnie");
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
