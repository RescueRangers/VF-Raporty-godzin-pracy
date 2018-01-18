﻿using System.IO;
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
                var znakiDoWyciecia = new char[2] { ' ', '\n' };

                using (var excel = new ExcelPackage(new FileInfo(template)))
                {
                    var dlugoscRaportu = raport.GetNaglowki().Count;
                    var wysokoscRaportu = pracownik.GetDni().Count;

                    //Nazwa pracownika w komorce A1, pozniej jest merge tej komorki na cala dlugosc raportu
                    excel.Workbook.Worksheets[1].Name = pracownik.NazwaPracownika();

                    var arkusz = excel.Workbook.Worksheets[1];
                    arkusz.Cells[1, 1].Value = pracownik.NazwaPracownika();
                    arkusz.Cells[1, 1, 1, dlugoscRaportu + 1].Merge = true;

                    var naglowekIndeks = 0;
                    var godziny = 0;
                    

                    arkusz.Cells[2, 1].Value = "Data";

                    //Zapelnianie raportu naglowkami
                    foreach (var naglowek in raport.GetNaglowki())
                    {
                        if (naglowek.Nazwa.ToLower() == "godziny pracy")
                        {
                            godziny = naglowekIndeks +2;
                        }

                        var tekstNaglowka = "";
                        var dlugoscTekstu = 0;

                        var slowa = naglowek.Nazwa.Split(' ');

                        //Wstawiam nowe linie jezeli slowo ma 5 lub wiecej liter
                        foreach (var slowo in slowa)
                        {
                            dlugoscTekstu = +slowo.Length + 1;
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
                    foreach (var dzien in pracownik.GetDni())
                    {
                        var godzinaIndeks = 0;
                        arkusz.Cells[3 + dzienIndeks, 1].Value = dzien.Date;
                        arkusz.Cells[3 + dzienIndeks, 1].Style.Numberformat.Format = "dd-mm-yyyy";

                        //Wstawianie godzin do raportu
                        foreach (var godzina in dzien.GetGodziny())
                        {
                            arkusz.Cells[3 + dzienIndeks, 2 + godzinaIndeks].Value = godzina;
                            godzinaIndeks++;
                        }
                        dzienIndeks++;
                    }

                    //Wstawianie dodatkowej kolumny w ktorej beda liczone pprawne godziny
                    arkusz.InsertColumn(godziny+1, 1);
                    arkusz.Cells[2, godziny+1].Value = "Godziny\npracy";

                    arkusz.Cells[2, dlugoscRaportu + 2].Value = "Razem";

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

                    excel.SaveAs(new FileInfo(nazwaPliku));
                }
            }
        }
    }
}