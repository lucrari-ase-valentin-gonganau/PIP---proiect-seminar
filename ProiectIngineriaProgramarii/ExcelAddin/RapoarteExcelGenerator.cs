using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public void GenereazaRaportComplet(List<Factura> facturi, string caleFisier)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheetProduse = null;
            Excel.Worksheet worksheetClienti = null;
            Excel.Worksheet worksheetVanzari = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                workbook = excelApp.Workbooks.Add();

                worksheetProduse = (Excel.Worksheet)workbook.Worksheets.Add();
                worksheetProduse.Name = "Top Produse";

                worksheetClienti = (Excel.Worksheet)workbook.Worksheets.Add();
                worksheetClienti.Name = "Top Clienti";

                worksheetVanzari = (Excel.Worksheet)workbook.Worksheets[1];
                worksheetVanzari.Name = "Vanzari Lunare";

                var topProduse = CalculeazaTopProduse(facturi);
                CreazaSheetTopProduse(worksheetProduse, topProduse);

                var topClienti = CalculeazaTopClienti(facturi);
                CreazaSheetTopClienti(worksheetClienti, topClienti);

                var vanzariLunare = CalculeazaVanzariPeLuni(facturi);
                CreazaSheetVanzariLunare(worksheetVanzari, vanzariLunare);

                workbook.SaveAs(caleFisier);
                workbook.Close();
                excelApp.Quit();

                if (worksheetProduse != null) Marshal.ReleaseComObject(worksheetProduse);
                if (worksheetClienti != null) Marshal.ReleaseComObject(worksheetClienti);
                if (worksheetVanzari != null) Marshal.ReleaseComObject(worksheetVanzari);
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

                throw new Exception($"Eroare la generarea raportului complet: {ex.Message}", ex);
            }
        }

        private Dictionary<string, decimal> CalculeazaTopProduse(List<Factura> facturi)
        {
            var produseCantitati = new Dictionary<string, decimal>();

            foreach (var factura in facturi)
            {
                foreach (var item in factura.Itemi)
                {
                    if (!produseCantitati.ContainsKey(item.NumeProdus))
                    {
                        produseCantitati[item.NumeProdus] = 0;
                    }
                    produseCantitati[item.NumeProdus] += item.Cantitate;
                }
            }

            return produseCantitati.OrderByDescending(x => x.Value)
                                   .Take(10)
                                   .ToDictionary(x => x.Key, x => x.Value);
        }

        private Dictionary<string, decimal> CalculeazaTopClienti(List<Factura> facturi)
        {
            var clientiValori = new Dictionary<string, decimal>();

            foreach (var factura in facturi)
            {
                if (factura.Client != null)
                {
                    string numeClient = factura.Client.NumeComplet;
                    if (!clientiValori.ContainsKey(numeClient))
                    {
                        clientiValori[numeClient] = 0;
                    }
                    clientiValori[numeClient] += factura.Total;
                }
            }

            return clientiValori.OrderByDescending(x => x.Value)
                                .Take(10)
                                .ToDictionary(x => x.Key, x => x.Value);
        }

        private Dictionary<string, decimal> CalculeazaVanzariPeLuni(List<Factura> facturi)
        {
            var vanzariLunare = new Dictionary<string, decimal>();

            foreach (var factura in facturi)
            {
                string lunaAn = factura.DataEmitere.ToString("yyyy-MM");
                if (!vanzariLunare.ContainsKey(lunaAn))
                {
                    vanzariLunare[lunaAn] = 0;
                }
                vanzariLunare[lunaAn] += factura.Total;
            }

            return vanzariLunare.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        }

        private void CreazaSheetTopProduse(Excel.Worksheet worksheet, Dictionary<string, decimal> topProduse)
        {
            worksheet.Cells[1, 1] = "TOP 10 PRODUSE CELE MAI VÂNDUTE";
            worksheet.Range["A1", "C1"].Merge();
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;
            worksheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int startRow = 3;
            worksheet.Cells[startRow, 1] = "Poziție";
            worksheet.Cells[startRow, 2] = "Nume Produs";
            worksheet.Cells[startRow, 3] = "Cantitate Vândută";

            Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"C{startRow}"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbLightBlue;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int currentRow = startRow + 1;
            int pozitie = 1;
            foreach (var produs in topProduse)
            {
                worksheet.Cells[currentRow, 1] = pozitie;
                worksheet.Cells[currentRow, 2] = produs.Key;
                worksheet.Cells[currentRow, 3] = produs.Value;
                pozitie++;
                currentRow++;
            }

            Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"C{currentRow - 1}"];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 10;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 35;
            ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 20;
        }

        private void CreazaSheetTopClienti(Excel.Worksheet worksheet, Dictionary<string, decimal> topClienti)
        {
            worksheet.Cells[1, 1] = "TOP 10 CLIENȚI";
            worksheet.Range["A1", "C1"].Merge();
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;
            worksheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int startRow = 3;
            worksheet.Cells[startRow, 1] = "Poziție";
            worksheet.Cells[startRow, 2] = "Nume Client";
            worksheet.Cells[startRow, 3] = "Valoare Totală (RON)";

            Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"C{startRow}"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGreen;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int currentRow = startRow + 1;
            int pozitie = 1;
            foreach (var client in topClienti)
            {
                worksheet.Cells[currentRow, 1] = pozitie;
                worksheet.Cells[currentRow, 2] = client.Key;
                worksheet.Cells[currentRow, 3] = client.Value;
                pozitie++;
                currentRow++;
            }

            Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"C{currentRow - 1}"];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range valuesRange = worksheet.Range[$"C{startRow + 1}", $"C{currentRow - 1}"];
            valuesRange.NumberFormat = "#,##0.00";

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 10;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 35;
            ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 20;
        }

        private void CreazaSheetVanzariLunare(Excel.Worksheet worksheet, Dictionary<string, decimal> vanzariLunare)
        {
            worksheet.Cells[1, 1] = "VÂNZĂRI PE LUNI";
            worksheet.Range["A1", "B1"].Merge();
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;
            worksheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int startRow = 3;
            worksheet.Cells[startRow, 1] = "Luna";
            worksheet.Cells[startRow, 2] = "Vânzări (RON)";

            Excel.Range headerRange = worksheet.Range[$"A{startRow}", $"B{startRow}"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbLightYellow;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int currentRow = startRow + 1;
            foreach (var luna in vanzariLunare)
            {
                worksheet.Cells[currentRow, 1] = luna.Key;
                worksheet.Cells[currentRow, 2] = luna.Value;
                currentRow++;
            }

            Excel.Range tableRange = worksheet.Range[$"A{startRow}", $"B{currentRow - 1}"];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range valuesRange = worksheet.Range[$"B{startRow + 1}", $"B{currentRow - 1}"];
            valuesRange.NumberFormat = "#,##0.00";

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 15;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 20;

            if (vanzariLunare.Count > 0)
            {
                Excel.Range chartRange = worksheet.Range[$"A{startRow}", $"B{currentRow - 1}"];
                Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                Excel.ChartObject chartObject = chartObjects.Add(480, 50, 500, 300);
                Excel.Chart chart = chartObject.Chart;

                chart.SetSourceData(chartRange);
                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Evoluția Vânzărilor pe Luni";
                chart.HasLegend = false;

                Excel.Axis xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
                xAxis.HasTitle = true;
                xAxis.AxisTitle.Text = "Luna";

                Excel.Axis yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                yAxis.HasTitle = true;
                yAxis.AxisTitle.Text = "Vânzări (RON)";
            }
        }
    }
}
