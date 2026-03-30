using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ProiectIngineriaProgramarii.Addin;
using ProiectIngineriaProgramarii.Data;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii
{
    public partial class ListaFacturiForm : Form
    {
        private StartForm _mainForm;
        private DatabaseManager _dbManager;
        private FacturaRepository _facturaRepository;
        private ClientRepository _clientRepository;
        private FacturaWordGenerator _wordGenerator;

        public ListaFacturiForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Lista Facturi";
            _mainForm = mainForm;
            InitializeRepositories();
            IncarcaFacturi();
        }

        public ListaFacturiForm()
        {
            InitializeComponent();
            this.Text = "Lista Facturi";
            InitializeRepositories();
            IncarcaFacturi();
        }

        private void InitializeRepositories()
        {
            _dbManager = new DatabaseManager();
            _clientRepository = new ClientRepository(_dbManager);
            _facturaRepository = new FacturaRepository(_dbManager, _clientRepository);
            _wordGenerator = new FacturaWordGenerator();
        }

        private void IncarcaFacturi()
        {
            try
            {
                var facturi = _facturaRepository.GetAll();
                
                var facturaDisplay = facturi.Select(f => new
                {
                    f.Id,
                    f.NumarFactura,
                    DataEmitere = f.DataEmitere.ToString("dd.MM.yyyy HH:mm"),
                    Client = f.Client?.NumeComplet ?? "N/A",
                    Subtotal = f.Subtotal,
                    TVA = f.TVA,
                    Total = f.Total,
                    f.Status,
                    f.Observatii
                }).ToList();

                dgvFacturi.DataSource = facturaDisplay;
                ConfigureazaDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la incarcarea facturilor: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigureazaDataGridView()
        {
            if (dgvFacturi.Columns.Count > 0)
            {
                // Change AutoSizeColumnsMode to allow manual width setting
                dgvFacturi.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                
                dgvFacturi.Columns["Id"].Visible = false;
                dgvFacturi.Columns["NumarFactura"].HeaderText = "Numar Factura";
                dgvFacturi.Columns["DataEmitere"].HeaderText = "Data Emitere";
                dgvFacturi.Columns["Client"].HeaderText = "Client";
                dgvFacturi.Columns["Subtotal"].HeaderText = "Subtotal (RON)";
                dgvFacturi.Columns["TVA"].HeaderText = "TVA (RON)";
                dgvFacturi.Columns["Total"].HeaderText = "Total (RON)";
                dgvFacturi.Columns["Status"].HeaderText = "Status";
                dgvFacturi.Columns["Observatii"].HeaderText = "Observatii";

                dgvFacturi.Columns["Subtotal"].DefaultCellStyle.Format = "N2";
                dgvFacturi.Columns["TVA"].DefaultCellStyle.Format = "N2";
                dgvFacturi.Columns["Total"].DefaultCellStyle.Format = "N2";

                dgvFacturi.Columns["NumarFactura"].Width = 150;
                dgvFacturi.Columns["DataEmitere"].Width = 130;
                dgvFacturi.Columns["Client"].Width = 150;
                dgvFacturi.Columns["Subtotal"].Width = 100;
                dgvFacturi.Columns["TVA"].Width = 100;
                dgvFacturi.Columns["Total"].Width = 100;
                dgvFacturi.Columns["Status"].Width = 80;
            }
        }

        private void btnGenereazaWord_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvFacturi.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Selecteaza o factura din lista!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedRow = dgvFacturi.SelectedRows[0];
                int facturaId = Convert.ToInt32(selectedRow.Cells["Id"].Value);

                var factura = _facturaRepository.GetById(facturaId);
                if (factura == null)
                {
                    MessageBox.Show("Factura nu a fost gasita!", "Eroare",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                this.Cursor = Cursors.WaitCursor;
                btnGenereazaWord.Enabled = false;
                btnGenereazaWord.Text = "Se genereaza factura...";
                Application.DoEvents();

                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string fileName = $"Factura_{factura.NumarFactura}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                string filePath = Path.Combine(documentsPath, fileName);

                _wordGenerator.GenereazaFactura(factura, filePath);
                _wordGenerator.DeschideFactura(filePath);

                btnGenereazaWord.Text = "Genereaza factura in Word";
                btnGenereazaWord.Enabled = true;
                this.Cursor = Cursors.Default;

                MessageBox.Show(
                    $"Factura a fost generata cu succes!\n\nLocatie: {filePath}",
                    "Succes",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnGenereazaWord.Text = "Genereaza factura in Word";
                btnGenereazaWord.Enabled = true;
                this.Cursor = Cursors.Default;

                MessageBox.Show($"Eroare la generarea facturii: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
