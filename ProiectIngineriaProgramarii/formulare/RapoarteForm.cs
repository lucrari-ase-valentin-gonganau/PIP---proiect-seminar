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
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
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

                // Sterge Sheet1 implicit (cel gol)
                try
                {
                    Excel.Worksheet sheet1 = null;
                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        if (ws.Name == "Sheet1" || ws.Name == "Foaie1")
                        {
                            sheet1 = ws;
                            break;
                        }
                    }
                    if (sheet1 != null)
                    {
                        excelApp.DisplayAlerts = false;
                        sheet1.Delete();
                        excelApp.DisplayAlerts = false;
                    }
                }
                catch { }

                // Adauga buton pentru editare date (deschide form C#)
                AdaugaButonEditareDate(worksheetVanzari, workbook);

                // Formatare log sheet
                ((Excel.Range)logSheet.Columns["A:A"]).ColumnWidth = 12;
                ((Excel.Range)logSheet.Columns["B:B"]).ColumnWidth = 18;
                ((Excel.Range)logSheet.Columns["C:C"]).ColumnWidth = 40;
                ((Excel.Range)logSheet.Columns["D:D"]).ColumnWidth = 15;
                ((Excel.Range)logSheet.Columns["F:F"]).ColumnWidth = 8;
                ((Excel.Range)logSheet.Columns["G:G"]).ColumnWidth = 16;
                ((Excel.Range)logSheet.Columns["H:H"]).ColumnWidth = 35;

                // ACUM ATASAM EVENIMENTELE - DUPA CE TOATE SHEET-URILE SUNT GATA

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



                // Facem Excel vizibil si activ
                excelApp.ScreenUpdating = true;
                excelApp.Visible = true;
                excelApp.DisplayAlerts = true;
                logSheet.Activate();

                MessageBox.Show(
                    "Raportul interactiv a fost generat cu succes!",
                    "Raport Interactiv cu Logging Activ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                if (excelApp != null)
                {
                    try 
                    { 
                        excelApp.ScreenUpdating = true;
                        excelApp.DisplayAlerts = true;
                    } 
                    catch { }
                }

                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                }

                MessageBox.Show(
                    $"Eroare COM la comunicarea cu Excel:\n{comEx.Message}\n\nAsigurati-va ca Excel este instalat corect.",
                    "Eroare Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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

        private static Excel.Workbook _workbookActiv = null;

        private void AdaugaButonEditareDate(Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            try
            {
                _workbookActiv = workbook;

                // Adauga modul VBA cu macro pentru editare
                string moduleName = "ModulEditareVanzari";
                string macroCode = @"
Sub EditareVanzariInteractiv()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Vanzari Lunare"")

    ' Afla cate luni sunt
    Dim ultimRand As Long
    ultimRand = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If ultimRand < 4 Then
        MsgBox ""Nu exista date de editat!"", vbExclamation
        Exit Sub
    End If

    ' Construieste lista de luni
    Dim listLuni As String
    listLuni = ""Selectati luna de editat:"" & vbCrLf & vbCrLf
    Dim i As Long
    For i = 4 To ultimRand
        If ws.Cells(i, 1).Value <> """" Then
            listLuni = listLuni & (i - 3) & "". "" & ws.Cells(i, 1).Value & "" - "" & Format(ws.Cells(i, 2).Value, ""#,##0.00"") & "" RON"" & vbCrLf
        End If
    Next i

    ' Cere utilizatorului sa selecteze linia
    Dim inputLinie As String
    inputLinie = InputBox(listLuni & vbCrLf & ""Introduceti numarul liniei (1-"" & (ultimRand - 3) & ""):"", ""Editare Vanzari"")

    If inputLinie = """" Then Exit Sub

    If Not IsNumeric(inputLinie) Then
        MsgBox ""Introduceti un numar valid!"", vbExclamation
        Exit Sub
    End If

    Dim linie As Long
    linie = CLng(inputLinie) + 3

    If linie < 4 Or linie > ultimRand Then
        MsgBox ""Numar linie invalid!"", vbExclamation
        Exit Sub
    End If

    ' Afiseaza luna selectata si cere noua valoare
    Dim lunaSelectata As String
    Dim valoareVeche As Double
    lunaSelectata = ws.Cells(linie, 1).Value
    valoareVeche = ws.Cells(linie, 2).Value

    Dim nouaValoare As String
    nouaValoare = InputBox(""Luna: "" & lunaSelectata & vbCrLf & vbCrLf & _
                           ""Valoare actuala: "" & Format(valoareVeche, ""#,##0.00"") & "" RON"" & vbCrLf & vbCrLf & _
                           ""Introduceti noua valoare:"", _
                           ""Editare Valoare Vanzari"", _
                           Format(valoareVeche, ""0.00""))

    If nouaValoare = """" Then Exit Sub

    If Not IsNumeric(nouaValoare) Then
        MsgBox ""Introduceti o valoare numerica valida!"", vbExclamation
        Exit Sub
    End If

    ' Actualizeaza valoarea
    ws.Cells(linie, 2).Value = CDbl(nouaValoare)

    ' Refresh chart (graficul se actualizeaza automat)
    MsgBox ""Valoarea pentru "" & lunaSelectata & "" a fost actualizata la "" & Format(CDbl(nouaValoare), ""#,##0.00"") & "" RON!"", vbInformation, ""Succes""
End Sub
";

                // Adauga modulul VBA direct prin Workbook
                var vbComp = workbook.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                vbComp.Name = moduleName;
                vbComp.CodeModule.AddFromString(macroCode);

                // Adauga buton care apeleaza macro-ul
                Excel.Buttons buttons = (Excel.Buttons)worksheet.Buttons(Type.Missing);
                Excel.Button btn = (Excel.Button)buttons.Add(480, 380, 150, 35);
                btn.Caption = "Editeaza Date";
                btn.OnAction = moduleName + ".EditareVanzariInteractiv";

                // Font buton
                btn.Font.Size = 11;
                btn.Font.Bold = true;

                // Adauga mesaj informativ
                Excel.Range celulaMesaj = (Excel.Range)worksheet.Cells[24, 1];
                celulaMesaj.Value = "Apasati butonul 'Editeaza Date' pentru a modifica valorile vanzarilor";
                celulaMesaj.Font.Italic = true;
                celulaMesaj.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            }
            catch (Exception ex)
            {
                // Daca nu merge cu VBA (nu e activat Trust Access), pune doar mesaj
                try
                {
                    Excel.Range celulaMesaj = (Excel.Range)worksheet.Cells[24, 1];
                    celulaMesaj.Value = "Pentru editare: modificati direct valorile din coloana B (Vanzari). Graficul se actualizeaza automat.";
                    celulaMesaj.Font.Italic = true;
                    celulaMesaj.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);
                }
                catch { }

                System.Diagnostics.Debug.WriteLine($"Nu s-a putut adauga macro VBA: {ex.Message}");
            }
        }

        private void CreazaUserFormEditareVanzari(Excel.Application excelApp, Excel.Workbook workbook, Excel.Worksheet worksheet)
        {
            // METODA VECHE - NU MAI E FOLOSITA
            // Inlocuita cu AdaugaButonEditareDate care foloseste doar Excel Interop fara VBA
        }
    }
}
