using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ProiectIngineriaProgramarii.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProiectIngineriaProgramarii.ExcelAddin
{
    public class RapoarteExcelGenerator
    {
        public void GenereazaRaportVanzari(List<Factura> facturi, string caleFisier, DateTime dataStart, DateTime dataSfarsit)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Raport Vanzari";

                AdaugaAntetRaport(worksheet, dataStart, dataSfarsit);
                AdaugaTabelVanzari(worksheet, facturi);
                FormatareaTabelului(worksheet, facturi.Count);
                AdaugaTotalizatoare(worksheet, facturi, facturi.Count + 5);

                workbook.SaveAs(caleFisier);
                workbook.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                throw new Exception($"Eroare la generarea raportului: {ex.Message}", ex);
            }
        }

        private void AdaugaAntetRaport(Excel.Worksheet worksheet, DateTime dataStart, DateTime dataSfarsit)
        {
            worksheet.Cells[1, 1] = "RAPORT VÂNZĂRI";
            worksheet.Range["A1", "F1"].Merge();
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;
            worksheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            worksheet.Cells[2, 1] = $"Perioada: {dataStart:dd.MM.yyyy} - {dataSfarsit:dd.MM.yyyy}";
            worksheet.Range["A2", "F2"].Merge();
            worksheet.Range["A2"].Font.Size = 11;
            worksheet.Range["A2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        private void AdaugaTabelVanzari(Excel.Worksheet worksheet, List<Factura> facturi)
        {
            int startRow = 4;

            worksheet.Cells[startRow, 1] = "Nr. Factură";
            worksheet.Cells[startRow, 2] = "Data";
            worksheet.Cells[startRow, 3] = "Client";
            worksheet.Cells[startRow, 4] = "Subtotal (RON)";
            worksheet.Cells[startRow, 5] = "TVA (RON)";
            worksheet.Cells[startRow, 6] = "Total (RON)";

            Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"F{startRow}"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int currentRow = startRow + 1;
            foreach (var factura in facturi)
            {
                worksheet.Cells[currentRow, 1] = factura.NumarFactura;
                worksheet.Cells[currentRow, 2] = factura.DataEmitere.ToString("dd.MM.yyyy");
                worksheet.Cells[currentRow, 3] = factura.Client != null 
                    ? $"{factura.Client.Nume} {factura.Client.Prenume}" 
                    : "N/A";
                worksheet.Cells[currentRow, 4] = factura.Subtotal;
                worksheet.Cells[currentRow, 5] = factura.TVA;
                worksheet.Cells[currentRow, 6] = factura.Total;

                currentRow++;
            }
        }

        private void FormatareaTabelului(Excel.Worksheet worksheet, int numarFacturi)
        {
            int startRow = 4;
            int endRow = startRow + numarFacturi;

            Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"F{endRow}"];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range dataRange = worksheet.Range[$"D{startRow + 1}", $"F{endRow}"];
            dataRange.NumberFormat = "#,##0.00";

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 18;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 12;
            ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 25;
            ((Excel.Range)worksheet.Columns["D:D"]).ColumnWidth = 15;
            ((Excel.Range)worksheet.Columns["E:E"]).ColumnWidth = 12;
            ((Excel.Range)worksheet.Columns["F:F"]).ColumnWidth = 15;
        }

        private void AdaugaTotalizatoare(Excel.Worksheet worksheet, List<Factura> facturi, int startRow)
        {
            decimal totalSubtotal = 0;
            decimal totalTVA = 0;
            decimal totalGeneral = 0;

            foreach (var factura in facturi)
            {
                totalSubtotal += factura.Subtotal;
                totalTVA += factura.TVA;
                totalGeneral += factura.Total;
            }

            worksheet.Cells[startRow, 3] = "TOTAL GENERAL:";
            worksheet.Cells[startRow, 4] = totalSubtotal;
            worksheet.Cells[startRow, 5] = totalTVA;
            worksheet.Cells[startRow, 6] = totalGeneral;

            Excel.Range totalRange = worksheet.Range[$"C{startRow}", $"F{startRow}"];
            totalRange.Font.Bold = true;
            totalRange.Interior.Color = Excel.XlRgbColor.rgbYellow;

            Excel.Range totalValuesRange = worksheet.Range[$"D{startRow}", $"F{startRow}"];
            totalValuesRange.NumberFormat = "#,##0.00";

            worksheet.Cells[startRow + 2, 1] = $"Total facturi: {facturi.Count}";
            worksheet.Range[$"A{startRow + 2}"].Font.Bold = true;
        }

        public void GenereazaRaportProduse(List<Produs> produse, string caleFisier)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Raport Produse";

                worksheet.Cells[1, 1] = "RAPORT PRODUSE";
                worksheet.Range["A1", "E1"].Merge();
                worksheet.Range["A1"].Font.Size = 16;
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int startRow = 3;
                worksheet.Cells[startRow, 1] = "ID";
                worksheet.Cells[startRow, 2] = "Nume Produs";
                worksheet.Cells[startRow, 3] = "Preț (RON)";
                worksheet.Cells[startRow, 4] = "Stoc Disponibil";
                worksheet.Cells[startRow, 5] = "Valoare Stoc (RON)";

                Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"E{startRow}"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int currentRow = startRow + 1;
                decimal valoareTotalaStoc = 0;

                foreach (var produs in produse)
                {
                    decimal valoareStoc = produs.Pret * produs.StocDisponibil;
                    valoareTotalaStoc += valoareStoc;

                    worksheet.Cells[currentRow, 1] = produs.Id;
                    worksheet.Cells[currentRow, 2] = produs.Nume;
                    worksheet.Cells[currentRow, 3] = produs.Pret;
                    worksheet.Cells[currentRow, 4] = produs.StocDisponibil;
                    worksheet.Cells[currentRow, 5] = valoareStoc;

                    currentRow++;
                }

                Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"E{currentRow - 1}"];
                tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                Excel.Range priceRange = worksheet.Range[$"C{startRow + 1}", $"C{currentRow - 1}"];
                priceRange.NumberFormat = "#,##0.00";

                Excel.Range valueRange = worksheet.Range[$"E{startRow + 1}", $"E{currentRow - 1}"];
                valueRange.NumberFormat = "#,##0.00";

                worksheet.Cells[currentRow + 1, 4] = "TOTAL:";
                worksheet.Cells[currentRow + 1, 5] = valoareTotalaStoc;
                Excel.Range totalRange = worksheet.Range[$"D{currentRow + 1}", $"E{currentRow + 1}"];
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Excel.XlRgbColor.rgbYellow;
                ((Excel.Range)worksheet.Cells[currentRow + 1, 5]).NumberFormat = "#,##0.00";

                ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 8;
                ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 30;
                ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["D:D"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["E:E"]).ColumnWidth = 18;

                workbook.SaveAs(caleFisier);
                workbook.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                throw new Exception($"Eroare la generarea raportului produse: {ex.Message}", ex);
            }
        }

        public void GenereazaRaportClienti(List<Client> clienti, string caleFisier)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Raport Clienti";

                worksheet.Cells[1, 1] = "RAPORT CLIENTI";
                var range1 = worksheet.Range["A1", "F1"];
                range1.Merge();
                range1.Font.Size = 16;
                range1.Font.Bold = true;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int startRow = 3;
                worksheet.Cells[startRow, 1] = "ID";
                worksheet.Cells[startRow, 2] = "Nume";
                worksheet.Cells[startRow, 3] = "Prenume";
                worksheet.Cells[startRow, 4] = "Email";
                worksheet.Cells[startRow, 5] = "Telefon";
                worksheet.Cells[startRow, 6] = "Data Inregistrare";

                Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"F{startRow}"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int currentRow = startRow + 1;
                foreach (var client in clienti)
                {
                    worksheet.Cells[currentRow, 1] = client.Id;
                    worksheet.Cells[currentRow, 2] = client.Nume ?? "";
                    worksheet.Cells[currentRow, 3] = client.Prenume ?? "";
                    worksheet.Cells[currentRow, 4] = client.Email ?? "";
                    worksheet.Cells[currentRow, 5] = client.Telefon ?? "";
                    worksheet.Cells[currentRow, 6] = client.DataInregistrare.ToString("dd.MM.yyyy");

                    currentRow++;
                }

                Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"F{currentRow - 1}"];
                tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 8;
                ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 20;
                ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 20;
                ((Excel.Range)worksheet.Columns["D:D"]).ColumnWidth = 30;
                ((Excel.Range)worksheet.Columns["E:E"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["F:F"]).ColumnWidth = 18;

                worksheet.Cells[currentRow + 1, 1] = $"Total clienti: {clienti.Count}";
                var boldRange = worksheet.Range[$"A{currentRow + 1}"];
                boldRange.Font.Bold = true;

                workbook.SaveAs(caleFisier);
                workbook.Close(false);
                excelApp.Quit();

                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch { }

                throw new Exception($"Eroare la generarea raportului clienti: {ex.Message}", ex);
            }
        }

        public void DeschideRaportClientiDirect(List<Client> clienti)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Raport Clienti";

                worksheet.Cells[1, 1] = "RAPORT CLIENTI";
                var range1 = worksheet.Range["A1", "F1"];
                range1.Merge();
                range1.Font.Size = 16;
                range1.Font.Bold = true;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int startRow = 3;
                worksheet.Cells[startRow, 1] = "ID";
                worksheet.Cells[startRow, 2] = "Nume";
                worksheet.Cells[startRow, 3] = "Prenume";
                worksheet.Cells[startRow, 4] = "Email";
                worksheet.Cells[startRow, 5] = "Telefon";
                worksheet.Cells[startRow, 6] = "Data Inregistrare";

                Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"F{startRow}"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int currentRow = startRow + 1;
                foreach (var client in clienti)
                {
                    worksheet.Cells[currentRow, 1] = client.Id;
                    worksheet.Cells[currentRow, 2] = client.Nume ?? "";
                    worksheet.Cells[currentRow, 3] = client.Prenume ?? "";
                    worksheet.Cells[currentRow, 4] = client.Email ?? "";
                    worksheet.Cells[currentRow, 5] = client.Telefon ?? "";
                    worksheet.Cells[currentRow, 6] = client.DataInregistrare.ToString("dd.MM.yyyy");

                    currentRow++;
                }

                Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"F{currentRow - 1}"];
                tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 8;
                ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 20;
                ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 20;
                ((Excel.Range)worksheet.Columns["D:D"]).ColumnWidth = 30;
                ((Excel.Range)worksheet.Columns["E:E"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["F:F"]).ColumnWidth = 18;

                worksheet.Cells[currentRow + 1, 1] = $"Total clienti: {clienti.Count}";
                var boldRange = worksheet.Range[$"A{currentRow + 1}"];
                boldRange.Font.Bold = true;
            }
            catch (Exception ex)
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { }

                throw new Exception($"Eroare la deschiderea raportului clienti: {ex.Message}", ex);
            }
        }

        public void DeschideRaportProduseDirect(List<Produs> produse)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Raport Produse";

                worksheet.Cells[1, 1] = "RAPORT PRODUSE";
                var range1 = worksheet.Range["A1", "E1"];
                range1.Merge();
                range1.Font.Size = 16;
                range1.Font.Bold = true;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int startRow = 3;
                worksheet.Cells[startRow, 1] = "ID";
                worksheet.Cells[startRow, 2] = "Nume Produs";
                worksheet.Cells[startRow, 3] = "Pret (RON)";
                worksheet.Cells[startRow, 4] = "Stoc Disponibil";
                worksheet.Cells[startRow, 5] = "Valoare Stoc (RON)";

                Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"E{startRow}"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int currentRow = startRow + 1;
                decimal valoareTotalaStoc = 0;

                foreach (var produs in produse)
                {
                    decimal valoareStoc = produs.Pret * produs.StocDisponibil;
                    valoareTotalaStoc += valoareStoc;

                    worksheet.Cells[currentRow, 1] = produs.Id;
                    worksheet.Cells[currentRow, 2] = produs.Nume ?? "";
                    worksheet.Cells[currentRow, 3] = produs.Pret;
                    worksheet.Cells[currentRow, 4] = produs.StocDisponibil;
                    worksheet.Cells[currentRow, 5] = valoareStoc;

                    currentRow++;
                }

                Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"E{currentRow - 1}"];
                tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                Excel.Range priceRange = worksheet.Range[$"C{startRow + 1}", $"C{currentRow - 1}"];
                priceRange.NumberFormat = "#,##0.00";

                Excel.Range valueRange = worksheet.Range[$"E{startRow + 1}", $"E{currentRow - 1}"];
                valueRange.NumberFormat = "#,##0.00";

                worksheet.Cells[currentRow + 1, 4] = "TOTAL:";
                worksheet.Cells[currentRow + 1, 5] = valoareTotalaStoc;
                var totalRange = worksheet.Range[$"D{currentRow + 1}", $"E{currentRow + 1}"];
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Excel.XlRgbColor.rgbYellow;
                ((Excel.Range)worksheet.Cells[currentRow + 1, 5]).NumberFormat = "#,##0.00";

                ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 8;
                ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 30;
                ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["D:D"]).ColumnWidth = 15;
                ((Excel.Range)worksheet.Columns["E:E"]).ColumnWidth = 18;
            }
            catch (Exception ex)
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { }

                throw new Exception($"Eroare la deschiderea raportului produse: {ex.Message}", ex);
            }
        }

        public void DeschideRaport(string caleFisier)
        {
            if (!File.Exists(caleFisier))
            {
                throw new FileNotFoundException("Fisierul nu a fost gasit.", caleFisier);
            }

            Excel.Application excelApp = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.Workbooks.Open(caleFisier);
            }
            catch (Exception ex)
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                throw new Exception($"Eroare la deschiderea raportului: {ex.Message}", ex);
            }
        }
    }
}
