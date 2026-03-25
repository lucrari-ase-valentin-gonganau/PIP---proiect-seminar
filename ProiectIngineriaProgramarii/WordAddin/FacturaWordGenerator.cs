using System;
using System.IO;
using System.Runtime.InteropServices;
using ProiectIngineriaProgramarii.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ProiectIngineriaProgramarii.WordAddin
{
    public class FacturaWordGenerator
    {
        public void GenereazaFactura(Factura factura, string caleFisier)
        {
            Word.Application wordApp = null;
            Word.Document document = null;

            try
            {
                wordApp = new Word.Application();
                document = wordApp.Documents.Add();

                document.PageSetup.TopMargin = 40f;
                document.PageSetup.BottomMargin = 40f;
                document.PageSetup.LeftMargin = 40f;
                document.PageSetup.RightMargin = 40f;

                AdaugaAntet(document, factura);
                AdaugaDetaliiClient(document, factura);
                AdaugaTabelProduse(document, factura);
                AdaugaTotaluri(document, factura);
                AdaugaSubsol(document, factura);

                document.SaveAs2(caleFisier);
                document.Close();
                wordApp.Quit();

                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(wordApp);
            }
            catch (Exception ex)
            {
                if (document != null)
                {
                    document.Close(false);
                    Marshal.ReleaseComObject(document);
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                throw new Exception($"Eroare la generarea facturii: {ex.Message}", ex);
            }
        }

        private void AdaugaAntet(Word.Document document, Factura factura)
        {
            var paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = "FACTURA\n";
            paragraph.Range.Font.Size = 24;
            paragraph.Range.Font.Bold = 1;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Range.InsertParagraphAfter();

            paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = $"Nr. {factura.NumarFactura}\n";
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Bold = 1;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Range.InsertParagraphAfter();

            paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = $"Data: {factura.DataEmitere:dd.MM.yyyy}\n\n";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AdaugaDetaliiClient(Word.Document document, Factura factura)
        {
            var paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = "DETALII CLIENT\n";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.InsertParagraphAfter();

            if (factura.Client != null)
            {
                paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = $"Nume: {factura.Client.Nume} {factura.Client.Prenume}\n";
                paragraph.Range.Font.Size = 11;
                paragraph.Range.Font.Bold = 0;
                paragraph.Range.InsertParagraphAfter();

                if (!string.IsNullOrEmpty(factura.Client.Email))
                {
                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Text = $"Email: {factura.Client.Email}\n";
                    paragraph.Range.Font.Size = 11;
                    paragraph.Range.InsertParagraphAfter();
                }

                if (!string.IsNullOrEmpty(factura.Client.Telefon))
                {
                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Text = $"Telefon: {factura.Client.Telefon}\n";
                    paragraph.Range.Font.Size = 11;
                    paragraph.Range.InsertParagraphAfter();
                }

                if (!string.IsNullOrEmpty(factura.Client.Adresa))
                {
                    paragraph = document.Paragraphs.Add();
                    paragraph.Range.Text = $"Adresa: {factura.Client.Adresa}\n";
                    paragraph.Range.Font.Size = 11;
                    paragraph.Range.InsertParagraphAfter();
                }
            }

            paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = "\n";
            paragraph.Range.InsertParagraphAfter();
        }

        private void AdaugaTabelProduse(Word.Document document, Factura factura)
        {
            int numRows = factura.Itemi.Count + 1;
            int numCols = 5;

            var range = document.Paragraphs[document.Paragraphs.Count].Range;
            var table = document.Tables.Add(range, numRows, numCols);

            table.Borders.Enable = 1;
            table.Range.Font.Size = 10;

            table.Cell(1, 1).Range.Text = "Nr.";
            table.Cell(1, 2).Range.Text = "Produs";
            table.Cell(1, 3).Range.Text = "Cantitate";
            table.Cell(1, 4).Range.Text = "Preț Unitar (RON)";
            table.Cell(1, 5).Range.Text = "Subtotal (RON)";

            table.Rows[1].Range.Font.Bold = 1;
            table.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;

            for (int i = 0; i < factura.Itemi.Count; i++)
            {
                var item = factura.Itemi[i];
                int rowIndex = i + 2;

                table.Cell(rowIndex, 1).Range.Text = (i + 1).ToString();
                table.Cell(rowIndex, 2).Range.Text = item.NumeProdus;
                table.Cell(rowIndex, 3).Range.Text = item.Cantitate.ToString();
                table.Cell(rowIndex, 4).Range.Text = item.PretUnitar.ToString("N2");
                table.Cell(rowIndex, 5).Range.Text = item.Subtotal.ToString("N2");
            }

            table.Columns[1].Width = 30f;
            table.Columns[2].Width = 180f;
            table.Columns[3].Width = 60f;
            table.Columns[4].Width = 80f;
            table.Columns[5].Width = 80f;

            var paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = "\n";
            paragraph.Range.InsertParagraphAfter();
        }

        private void AdaugaTotaluri(Word.Document document, Factura factura)
        {
            var paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = $"Subtotal: {factura.Subtotal:N2} RON\n";
            paragraph.Range.Font.Size = 11;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            paragraph.Range.InsertParagraphAfter();

            paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = $"TVA (19%): {factura.TVA:N2} RON\n";
            paragraph.Range.Font.Size = 11;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            paragraph.Range.InsertParagraphAfter();

            paragraph = document.Paragraphs.Add();
            paragraph.Range.Text = $"TOTAL: {factura.Total:N2} RON\n\n";
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Bold = 1;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AdaugaSubsol(Word.Document document, Factura factura)
        {
            if (!string.IsNullOrEmpty(factura.Observatii))
            {
                var paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = "Observații:\n";
                paragraph.Range.Font.Size = 10;
                paragraph.Range.Font.Bold = 1;
                paragraph.Range.InsertParagraphAfter();

                paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = $"{factura.Observatii}\n\n";
                paragraph.Range.Font.Size = 10;
                paragraph.Range.Font.Bold = 0;
                paragraph.Range.InsertParagraphAfter();
            }

            var paragraphFinal = document.Paragraphs.Add();
            paragraphFinal.Range.Text = "Vă mulțumim pentru colaborare!\n";
            paragraphFinal.Range.Font.Size = 10;
            paragraphFinal.Range.Font.Italic = 1;
            paragraphFinal.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraphFinal.Range.InsertParagraphAfter();
        }

        public void DeschideFactura(string caleFisier)
        {
            if (!File.Exists(caleFisier))
            {
                throw new FileNotFoundException("Fișierul nu a fost găsit.", caleFisier);
            }

            Word.Application wordApp = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = true;
                wordApp.Documents.Open(caleFisier);
            }
            catch (Exception ex)
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                throw new Exception($"Eroare la deschiderea facturii: {ex.Message}", ex);
            }
        }
    }
}
