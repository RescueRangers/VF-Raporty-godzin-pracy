using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System;

namespace VF_Raporty_Godzin_Pracy
{
    public class ZapiszExcel
    {
        /// <summary>
        /// Zapiuje raporty wszystkich pracowników do oddzielnych plików
        /// </summary>
        /// <param name="raport"></param>
        public static void ZapiszDoExcel(Raport raport, string folderDoZapisu)
        {
            Zapisz(raport, raport.GetPracownicy(), folderDoZapisu);
        }
        /// <summary>
        /// Zapisuje wybranego pracownika do pliku
        /// </summary>
        /// <param name="raport"></param>
        /// <param name="indeksPracownika"></param>
        public static void ZapiszDoExcel(Raport raport, List<Pracowik> nazwaPracownika, string folderDoZapisu)
        {
            Zapisz(raport, nazwaPracownika, folderDoZapisu);
        }

        private static void Zapisz(Raport raport, List<Pracowik> nazwaPracownika, string folderDoZapisu)
        {
            if (raport == null)
            {
                throw new InvalidDataException("Niepoprawny raport.");
            }
            if (nazwaPracownika == null)
            {
                throw new InvalidDataException("Nie wybrano pracownika z listy");
            }
            foreach (var pracownik in nazwaPracownika)
            {
                var template = $@"{AppDomain.CurrentDomain.BaseDirectory}Assets\template.xlsx";
                var nazwaPliku = $@"{folderDoZapisu}\{pracownik.NazwaPracownika()}.xlsx";


                using (var excel = new ExcelPackage(new FileInfo(template)))
                {
                    excel.Workbook.Worksheets[1].Name = pracownik.NazwaPracownika();
                    var arkusz = excel.Workbook.Worksheets[1];
                    arkusz.Cells[1, 1].Value = pracownik.NazwaPracownika();
                    arkusz.Cells[1, 1, 1, raport.GetNaglowki().Count].Merge = true;
                    var naglowekIndeks = 0;
                    var godziny = 0;
                    var nadgodziny = 0;
                    arkusz.Cells[2, 1].Value = "Data";
                    foreach (var naglowek in raport.GetNaglowki())
                    {
                        if (naglowek.Nazwa.ToLower() == "godziny pracy")
                        {
                            godziny = naglowekIndeks +2;
                        }
                        if (naglowek.Nazwa.ToLower() == "nadgodziny 50%")
                        {
                            nadgodziny = naglowekIndeks +2;
                        }

                        arkusz.Cells[2, 2 + naglowekIndeks].Value = naglowek.Nazwa;
                        naglowekIndeks++;
                    }

                    var dzienIndeks = 0;
                    foreach (var dzien in pracownik.GetDni())
                    {
                        var godzinaIndeks = 0;
                        arkusz.Cells[3 + dzienIndeks, 1].Value = dzien.Date;
                        arkusz.Cells[3 + dzienIndeks, 1].Style.Numberformat.Format = "dd-mm-yyyy";
                        foreach (var godzina in dzien.GetGodziny())
                        {
                            arkusz.Cells[3 + dzienIndeks, 2 + godzinaIndeks].Value = godzina;
                            godzinaIndeks++;
                        }
                        dzienIndeks++;
                    }
                    arkusz.InsertColumn(godziny+1, 1);
                    arkusz.Cells[2, godziny].Value = "Godziny\npracy";
                    for (int i = 3; i < pracownik.GetDni().Count; i++)
                    {
                        arkusz.Cells[i, godziny + 1].FormulaR1C1 = $"round(R{i}C{godziny}-R{i}C{godziny+2},0)";
                    }
                    arkusz.Column(godziny).Hidden = true;

                    excel.SaveAs(new FileInfo(nazwaPliku));
                }
            }
        }
    }
}