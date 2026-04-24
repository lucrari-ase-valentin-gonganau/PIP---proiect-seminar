using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using ProiectIngineriaProgramarii.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ProiectIngineriaProgramarii.Addin
{
    public class FacturaWordGenerator
    {
        private string GetCaleaSablon()
        {
            string caleTemplates = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            string caleSablon = Path.Combine(caleTemplates, "FacturaSablon.docx");

            if (!File.Exists(caleSablon))
            {
                throw new FileNotFoundException(
                    $"Sablonul facturii nu a fost gasit.\n\nCale asteptata:\n{caleSablon}\n\nPlasati fisierul 'FacturaSablon.docx' in folderul Templates si reincercati.",
                    caleSablon
                );
            }

            return caleSablon;
        }

        public void GenereazaFactura(Factura factura, string caleFisier)
        {
            Word.Application wordApp = null;
            Word.Document document = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                string caleSablon = GetCaleaSablon(); // arunca FileNotFoundException daca lipseste
                document = wordApp.Documents.Open(caleSablon);

                InlocuiestePlaceholders(document, factura);

                document.SaveAs2(caleFisier);
                document.Close(false);
                wordApp.Quit(false);

                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(wordApp);
            }
            catch (FileNotFoundException ex)
            {
                if (wordApp != null) { try { wordApp.Quit(false); } catch { } Marshal.ReleaseComObject(wordApp); }

                System.Windows.Forms.MessageBox.Show(
                    ex.Message,
                    "Sablon lipsa",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
            catch (Exception ex)
            {
                if (document != null) { try { document.Close(false); } catch { } Marshal.ReleaseComObject(document); }
                if (wordApp != null) { try { wordApp.Quit(false); } catch { } Marshal.ReleaseComObject(wordApp); }

                throw new Exception($"Eroare la generarea facturii: {ex.Message}", ex);
            }
        }

        private void InlocuiestePlaceholders(Word.Document document, Factura factura)
        {
            // Placeholders simple
            InlocuiesteText(document, "{{NUMAR_FACTURA}}", factura.NumarFactura);
            InlocuiesteText(document, "{{DATA_EMITERE}}", factura.DataEmitere.ToString("dd.MM.yyyy"));

            if (factura.Client != null)
            {
                InlocuiesteText(document, "{{NUME_CLIENT}}", $"{factura.Client.Nume} {factura.Client.Prenume}");
                InlocuiesteText(document, "{{ADRESA_CLIENT}}", factura.Client.Adresa ?? "");
                InlocuiesteText(document, "{{EMAIL_CLIENT}}", factura.Client.Email ?? "");
                InlocuiesteText(document, "{{TELEFON_CLIENT}}", factura.Client.Telefon ?? "");
            }

            InlocuiesteText(document, "{{SUBTOTAL}}", factura.Subtotal.ToString("N2"));
            InlocuiesteText(document, "{{TVA}}", factura.TVA.ToString("N2"));
            InlocuiesteText(document, "{{TOTAL}}", factura.Total.ToString("N2"));
            InlocuiesteText(document, "{{OBSERVATII}}", factura.Observatii ?? "");

            // Tabel produse - cauta placeholder si il inlocuieste cu tabel
            InlocuiesteCuTabelProduse(document, factura);
        }

        private void InlocuiesteText(Word.Document document, string placeholder, string valoare)
        {
            Word.Find findObject = document.Content.Find;
            findObject.Text = placeholder;
            findObject.Replacement.Text = valoare;
            findObject.Forward = true;
            findObject.Wrap = Word.WdFindWrap.wdFindContinue;
            findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        private void InlocuiesteCuTabelProduse(Word.Document document, Factura factura)
        {
            Word.Range range = document.Content;
            range.Find.ClearFormatting();
            range.Find.Text = "{{TABEL_PRODUSE}}";
            range.Find.Forward = true;
            range.Find.Wrap = Word.WdFindWrap.wdFindStop;

            if (!range.Find.Execute())
                return;

            // range-ul e acum pozitionat pe placeholder — il stergem
            range.Delete();

            // cream tabelul exact in locul ramas
            int numRows = factura.Itemi.Count + 1;
            Word.Table table = document.Tables.Add(range, numRows, 6);

            table.Borders.Enable = 1;
            table.Range.Font.Size = 10;
            table.Range.Font.Name = "Arial";
            table.Range.Font.Bold = 0;

            // header - randul 1
            string[] headers = { "Nr.", "Denumire produs / serviciu", "U.M.", "Cantitate", "Pret unitar (RON)", "Valoare (RON)" };
            for (int c = 0; c < headers.Length; c++)
            {
                Word.Cell headerCell = table.Cell(1, c + 1);
                headerCell.Range.Text = headers[c];
                headerCell.Range.Font.Bold = 0;
                headerCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerCell.Shading.BackgroundPatternColor = (Word.WdColor)0xE8E8E8;
            }

            // produse - incep de la randul 2
            for (int i = 0; i < factura.Itemi.Count; i++)
            {
                var item = factura.Itemi[i];
                int r = i + 2;

                table.Cell(r, 1).Range.Text = (i + 1).ToString();
                table.Cell(r, 2).Range.Text = item.NumeProdus;
                table.Cell(r, 3).Range.Text = item.UnitateMasura ?? "buc";
                table.Cell(r, 4).Range.Text = item.Cantitate.ToString();
                table.Cell(r, 5).Range.Text = item.PretUnitar.ToString("N2");
                table.Cell(r, 6).Range.Text = item.Subtotal.ToString("N2");
            }

            // latimi coloane (in puncte, total ~455pt pentru A4 cu margini 2.5cm)
            table.Columns[1].Width = 30f;   // Nr.
            table.Columns[2].Width = 185f;  // Denumire
            table.Columns[3].Width = 40f;   // U.M.
            table.Columns[4].Width = 50f;   // Cantitate
            table.Columns[5].Width = 75f;   // Pret unitar
            table.Columns[6].Width = 75f;   // Valoare
        }

       

        public void DeschideFactura(string caleFisier)
        {
            if (!File.Exists(caleFisier))
            {
                throw new FileNotFoundException("Fisierul nu a fost gasit.", caleFisier);
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
