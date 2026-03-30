using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ProiectIngineriaProgramarii.Data;
using ProiectIngineriaProgramarii.ExcelAddin;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProiectIngineriaProgramarii
{
    public partial class RapoarteForm : Form
    {
        private StartForm _mainForm;
        private DatabaseManager _dbManager;
        private FacturaRepository _facturaRepository;
        private ClientRepository _clientRepository;
        private ProdusRepository _produsRepository;

        public RapoarteForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Rapoarte";
            _mainForm = mainForm;
            InitializeRepositories();
        }

        public RapoarteForm()
        {
            InitializeComponent();
            this.Text = "Rapoarte";
            InitializeRepositories();
        }

        private void InitializeRepositories()
        {
            _dbManager = new DatabaseManager();
            _clientRepository = new ClientRepository(_dbManager);
            _produsRepository = new ProdusRepository(_dbManager);
            _facturaRepository = new FacturaRepository(_dbManager, _clientRepository);
        }

        private void btnRaportInteractiv_Click(object sender, EventArgs e)
        {
            try
            {
                // Afiseaza loading cursor
                this.Cursor = Cursors.WaitCursor;
                btnRaportInteractiv.Enabled = false;
                btnRaportInteractiv.Text = "Se genereaza...";
                Application.DoEvents();

                var facturi = _facturaRepository.GetAll();

                if (facturi == null || facturi.Count == 0)
                {
                    MessageBox.Show("Nu exista facturi in baza de date pentru a genera raportul.",
                        "Informatie", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                GenereazaRaportCuEvenimente(facturi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la generarea raportului interactiv: {ex.Message}",
                    "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Restaureaza cursor si buton
                this.Cursor = Cursors.Default;
                btnRaportInteractiv.Enabled = true;
                btnRaportInteractiv.Text = "Vezi Graficul In Excel (Cu Evenimente)";
            }
        }

        private void GenereazaRaportCuEvenimente(List<ProiectIngineriaProgramarii.Models.Factura> facturi)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet logSheet = null;
            int logRow = 2;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.DisplayAlerts = true;
                excelApp.ScreenUpdating = false;

                workbook = excelApp.Workbooks.Add();

                // Creeaza sheet pentru log evenimente
                logSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                logSheet.Name = "Log Evenimente";
                logSheet.Move(workbook.Worksheets[1]);

                // Header pentru log
                logSheet.Cells[1, 1] = "JURNAL EVENIMENTE OFFICE";
                logSheet.Range["A1", "D1"].Merge();
                logSheet.Range["A1"].Font.Size = 16;
                logSheet.Range["A1"].Font.Bold = true;
                logSheet.Range["A1"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                logSheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Info despre evenimente MONITORIZATE - in partea dreapta sus
                logSheet.Cells[1, 6] = "EVENIMENTE MONITORIZATE:";
                ((Excel.Range)logSheet.Cells[1, 6]).Font.Bold = true;
                ((Excel.Range)logSheet.Cells[1, 6]).Font.Size = 12;
                ((Excel.Range)logSheet.Cells[1, 6]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);

                int infoRow = 2;
                logSheet.Cells[infoRow, 6] = "[OK]";
                logSheet.Cells[infoRow, 7] = "BeforeSave";
                logSheet.Cells[infoRow, 8] = "Se declanseaza inainte de salvare";
                ((Excel.Range)logSheet.Cells[infoRow, 6]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                infoRow++;
                logSheet.Cells[infoRow, 6] = "[OK]";
                logSheet.Cells[infoRow, 7] = "SheetActivate";
                logSheet.Cells[infoRow, 8] = "Se declanseaza la schimbare sheet";
                ((Excel.Range)logSheet.Cells[infoRow, 6]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                infoRow++;
                logSheet.Cells[infoRow, 6] = "[OK]";
                logSheet.Cells[infoRow, 7] = "SheetChange";
                logSheet.Cells[infoRow, 8] = "Se declanseaza la modificare celula";
                ((Excel.Range)logSheet.Cells[infoRow, 6]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                logRow = 3;
                logSheet.Cells[logRow, 1] = "Timestamp";
                logSheet.Cells[logRow, 2] = "Tip Eveniment";
                logSheet.Cells[logRow, 3] = "Detalii";
                logSheet.Cells[logRow, 4] = "Status";

                Excel.Range headerRange = logSheet.Range["A3", "D3"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                logRow = 4;

                // Adauga eveniment initial
                logSheet.Cells[logRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                logSheet.Cells[logRow, 2] = "WorkbookOpen";
                logSheet.Cells[logRow, 3] = $"Raport creat cu {facturi.Count} facturi";
                logSheet.Cells[logRow, 4] = "ACTIV";
                ((Excel.Range)logSheet.Cells[logRow, 4]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                logRow++;

                // TRATARE EVENIMENT OFFICE - BeforeSave
                workbook.BeforeSave += (bool SaveAsUI, ref bool Cancel) =>
                {
                    try
                    {
                        logSheet.Cells[logRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                        logSheet.Cells[logRow, 2] = "BeforeSave";
                        logSheet.Cells[logRow, 3] = SaveAsUI ? "SaveAs Dialog" : "Save Direct";

                        var result = MessageBox.Show(
                            $"Se vor salva {facturi.Count} facturi in raport.\n\nContinuati salvarea?",
                            "Confirmare Salvare Raport",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.No)
                        {
                            Cancel = true;
                            logSheet.Cells[logRow, 4] = "ANULAT";
                            ((Excel.Range)logSheet.Cells[logRow, 4]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            logSheet.Cells[logRow, 4] = "SALVAT";
                            ((Excel.Range)logSheet.Cells[logRow, 4]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        }
                        logRow++;
                    }
                    catch { }
                };

                // TRATARE EVENIMENT OFFICE - SheetActivate  
                workbook.SheetActivate += (object Sh) =>
                {
                    try
                    {
                        var sheet = (Excel.Worksheet)Sh;
                        if (sheet.Name != "Log Evenimente")
                        {
                            logSheet.Cells[logRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                            logSheet.Cells[logRow, 2] = "SheetActivate";
                            logSheet.Cells[logRow, 3] = $"Navigare catre: {sheet.Name}";
                            logSheet.Cells[logRow, 4] = "ACTIV";
                            ((Excel.Range)logSheet.Cells[logRow, 4]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            logRow++;
                        }
                    }
                    catch { }
                };

                // TRATARE EVENIMENT OFFICE - SheetChange
                workbook.SheetChange += (object Sh, Excel.Range Target) =>
                {
                    try
                    {
                        var sheet = (Excel.Worksheet)Sh;
                        if (sheet.Name != "Log Evenimente")
                        {
                            logSheet.Cells[logRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                            logSheet.Cells[logRow, 2] = "SheetChange";
                            logSheet.Cells[logRow, 3] = $"Sheet: {sheet.Name}, Celula: {Target.Address}";
                            logSheet.Cells[logRow, 4] = "MODIFICAT";
                            ((Excel.Range)logSheet.Cells[logRow, 4]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                            logRow++;
                        }
                    }
                    catch { }
                };

                // Creeaza celelalte sheet-uri in ordine corecta
                Excel.Worksheet worksheetVanzari = (Excel.Worksheet)workbook.Worksheets.Add();
                worksheetVanzari.Name = "Vanzari Lunare";

                Excel.Worksheet worksheetClienti = (Excel.Worksheet)workbook.Worksheets.Add();
                worksheetClienti.Name = "Top Clienti";

                Excel.Worksheet worksheetProduse = (Excel.Worksheet)workbook.Worksheets.Add();
                worksheetProduse.Name = "Top Produse";

                // Genereaza raportul complet
                var generator = new RapoarteExcelGenerator();

                var topProduse = CalculeazaTopProduse(facturi);
                CreazaSheetTopProduseSimplificat(worksheetProduse, topProduse);

                var topClienti = CalculeazaTopClienti(facturi);
                CreazaSheetTopClientiSimplificat(worksheetClienti, topClienti);

                var vanzariLunare = CalculeazaVanzariPeLuni(facturi);
                CreazaSheetVanzariLunareSimplificat(worksheetVanzari, vanzariLunare);

                // Formatare log sheet
                ((Excel.Range)logSheet.Columns["A:A"]).ColumnWidth = 12;
                ((Excel.Range)logSheet.Columns["B:B"]).ColumnWidth = 18;
                ((Excel.Range)logSheet.Columns["C:C"]).ColumnWidth = 40;
                ((Excel.Range)logSheet.Columns["D:D"]).ColumnWidth = 15;
                ((Excel.Range)logSheet.Columns["F:F"]).ColumnWidth = 8;
                ((Excel.Range)logSheet.Columns["G:G"]).ColumnWidth = 16;
                ((Excel.Range)logSheet.Columns["H:H"]).ColumnWidth = 35;

                excelApp.ScreenUpdating = true;
                logSheet.Activate();

                MessageBox.Show(
                    "Raportul interactiv a fost generat cu succes!",
                    "Raport Interactiv cu Logging Activ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                }
                throw new Exception($"Eroare la generarea raportului interactiv: {ex.Message}", ex);
            }
        }

        private Dictionary<string, decimal> CalculeazaTopProduse(List<ProiectIngineriaProgramarii.Models.Factura> facturi)
        {
            var produseCantitati = new Dictionary<string, decimal>();
            foreach (var factura in facturi)
            {
                foreach (var item in factura.Itemi)
                {
                    if (!produseCantitati.ContainsKey(item.NumeProdus))
                        produseCantitati[item.NumeProdus] = 0;
                    produseCantitati[item.NumeProdus] += item.Cantitate;
                }
            }
            return produseCantitati.OrderByDescending(x => x.Value).Take(10).ToDictionary(x => x.Key, x => x.Value);
        }

        private Dictionary<string, decimal> CalculeazaTopClienti(List<ProiectIngineriaProgramarii.Models.Factura> facturi)
        {
            var clientiValori = new Dictionary<string, decimal>();
            foreach (var factura in facturi)
            {
                if (factura.Client != null)
                {
                    string numeClient = factura.Client.NumeComplet;
                    if (!clientiValori.ContainsKey(numeClient))
                        clientiValori[numeClient] = 0;
                    clientiValori[numeClient] += factura.Total;
                }
            }
            return clientiValori.OrderByDescending(x => x.Value).Take(10).ToDictionary(x => x.Key, x => x.Value);
        }

        private Dictionary<string, decimal> CalculeazaVanzariPeLuni(List<ProiectIngineriaProgramarii.Models.Factura> facturi)
        {
            var vanzariLunare = new Dictionary<string, decimal>();
            foreach (var factura in facturi)
            {
                string lunaAn = factura.DataEmitere.ToString("yyyy-MM");
                if (!vanzariLunare.ContainsKey(lunaAn))
                    vanzariLunare[lunaAn] = 0;
                vanzariLunare[lunaAn] += factura.Total;
            }
            return vanzariLunare.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        }

        private void CreazaSheetTopProduseSimplificat(Excel.Worksheet worksheet, Dictionary<string, decimal> topProduse)
        {
            worksheet.Cells[1, 1] = "TOP 10 PRODUSE";
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;

            int row = 3;
            worksheet.Cells[row, 1] = "Pozitie";
            worksheet.Cells[row, 2] = "Produs";
            worksheet.Cells[row, 3] = "Cantitate";
            ((Excel.Range)worksheet.Range["A3", "C3"]).Font.Bold = true;
            ((Excel.Range)worksheet.Range["A3", "C3"]).Interior.Color = Excel.XlRgbColor.rgbLightBlue;

            row = 4;
            int pozitie = 1;
            foreach (var produs in topProduse)
            {
                worksheet.Cells[row, 1] = pozitie;
                worksheet.Cells[row, 2] = produs.Key;
                worksheet.Cells[row, 3] = produs.Value;
                pozitie++;
                row++;
            }

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 10;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 35;
            ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 15;
        }

        private void CreazaSheetTopClientiSimplificat(Excel.Worksheet worksheet, Dictionary<string, decimal> topClienti)
        {
            worksheet.Cells[1, 1] = "TOP 10 CLIENTI";
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;

            int row = 3;
            worksheet.Cells[row, 1] = "Pozitie";
            worksheet.Cells[row, 2] = "Client";
            worksheet.Cells[row, 3] = "Valoare (RON)";
            ((Excel.Range)worksheet.Range["A3", "C3"]).Font.Bold = true;
            ((Excel.Range)worksheet.Range["A3", "C3"]).Interior.Color = Excel.XlRgbColor.rgbLightGreen;

            row = 4;
            int pozitie = 1;
            foreach (var client in topClienti)
            {
                worksheet.Cells[row, 1] = pozitie;
                worksheet.Cells[row, 2] = client.Key;
                worksheet.Cells[row, 3] = client.Value;
                ((Excel.Range)worksheet.Cells[row, 3]).NumberFormat = "#,##0.00";
                pozitie++;
                row++;
            }

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 10;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 35;
            ((Excel.Range)worksheet.Columns["C:C"]).ColumnWidth = 15;
        }

        private void CreazaSheetVanzariLunareSimplificat(Excel.Worksheet worksheet, Dictionary<string, decimal> vanzariLunare)
        {
            worksheet.Cells[1, 1] = "VANZARI PE LUNI";
            worksheet.Range["A1"].Font.Size = 16;
            worksheet.Range["A1"].Font.Bold = true;

            int row = 3;
            worksheet.Cells[row, 1] = "Luna";
            worksheet.Cells[row, 2] = "Vanzari (RON)";
            ((Excel.Range)worksheet.Range["A3", "B3"]).Font.Bold = true;
            ((Excel.Range)worksheet.Range["A3", "B3"]).Interior.Color = Excel.XlRgbColor.rgbLightYellow;

            row = 4;
            foreach (var luna in vanzariLunare)
            {
                worksheet.Cells[row, 1] = luna.Key;
                worksheet.Cells[row, 2] = luna.Value;
                ((Excel.Range)worksheet.Cells[row, 2]).NumberFormat = "#,##0.00";
                row++;
            }

            ((Excel.Range)worksheet.Columns["A:A"]).ColumnWidth = 15;
            ((Excel.Range)worksheet.Columns["B:B"]).ColumnWidth = 20;

            // Adauga grafic (chart) pentru vanzari lunare
            if (vanzariLunare.Count > 0)
            {
                Excel.Range chartRange = worksheet.Range[$"A3", $"B{row - 1}"];
                Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                Excel.ChartObject chartObject = chartObjects.Add(480, 50, 500, 300);
                Excel.Chart chart = chartObject.Chart;

                chart.SetSourceData(chartRange);
                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Evolutia Vanzarilor pe Luni";
                chart.HasLegend = false;

                Excel.Axis xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
                xAxis.HasTitle = true;
                xAxis.AxisTitle.Text = "Luna";

                Excel.Axis yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                yAxis.HasTitle = true;
                yAxis.AxisTitle.Text = "Vanzari (RON)";
            }
        }
    }
}
