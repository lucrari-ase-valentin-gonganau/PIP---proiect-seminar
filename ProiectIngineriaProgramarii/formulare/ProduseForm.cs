using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using ProiectIngineriaProgramarii.Data;
using ProiectIngineriaProgramarii.Models;
using ProiectIngineriaProgramarii.ExcelAddin;

namespace ProiectIngineriaProgramarii
{
    public partial class ProduseForm : Form
    {
        private StartForm _mainForm;
        private DatabaseManager _dbManager;
        private ProdusRepository _produsRepository;
        private RapoarteExcelGenerator _excelGenerator;

        public ProduseForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Gestiune Produse";
            _mainForm = mainForm;
            _dbManager = new DatabaseManager();
            _produsRepository = new ProdusRepository(_dbManager);
            _excelGenerator = new RapoarteExcelGenerator();

            IncarcaProduse();
        }

        public ProduseForm()
        {
            InitializeComponent();
            this.Text = "Gestiune Produse";
            _dbManager = new DatabaseManager();
            _produsRepository = new ProdusRepository(_dbManager);
            _excelGenerator = new RapoarteExcelGenerator();
            IncarcaProduse();
        }

        private void IncarcaProduse()
        {
            try
            {
                var produse = _produsRepository.GetAll();
                dgvProduse.DataSource = produse;

                if (dgvProduse.Columns.Count > 0)
                {
                    dgvProduse.Columns["Id"].HeaderText = "ID";
                    dgvProduse.Columns["Nume"].HeaderText = "Nume";
                    dgvProduse.Columns["Descriere"].HeaderText = "Descriere";
                    dgvProduse.Columns["Pret"].HeaderText = "Pret (RON)";
                    dgvProduse.Columns["StocDisponibil"].HeaderText = "Stoc";
                    dgvProduse.Columns["DataAdaugare"].HeaderText = "Data Adaugare";

                    dgvProduse.Columns["Pret"].DefaultCellStyle.Format = "N2";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la incarcarea produselor: {ex.Message}", "Eroare", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdauga_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtNume.Text))
                {
                    MessageBox.Show("Numele produsului este obligatoriu!", "Atentie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!decimal.TryParse(txtPret.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal pret))
                {
                    MessageBox.Show("Pretul trebuie sa fie un numar valid!", "Atentie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!int.TryParse(txtStoc.Text, out int stoc) || stoc < 0)
                {
                    MessageBox.Show("Stocul trebuie sa fie un numar intreg pozitiv!", "Atentie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var produs = new Produs
                {
                    Nume = txtNume.Text.Trim(),
                    Descriere = txtDescriere.Text.Trim(),
                    Pret = pret,
                    StocDisponibil = stoc,
                    DataAdaugare = DateTime.Now
                };

                _produsRepository.Add(produs);
                MessageBox.Show("Produs adaugat cu succes!", "Succes", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                CurataFormular();
                IncarcaProduse();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la adaugarea produsului: {ex.Message}", "Eroare", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVizualizeaza_Click(object sender, EventArgs e)
        {
            try
            {
                var produse = _produsRepository.GetAll();

                if (produse.Count == 0)
                {
                    MessageBox.Show("Nu exista produse in baza de date!", "Informatie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Afișează loader
                this.Cursor = Cursors.WaitCursor;
                btnVizualizeaza.Enabled = false;
                btnVizualizeaza.Text = "Se deschide Excel...";
                Application.DoEvents();

                // Deschide Excel
                _excelGenerator.DeschideRaportProduseDirect(produse);

                // Resetează UI
                btnVizualizeaza.Text = "Vizualizeaza in Excel";
                btnVizualizeaza.Enabled = true;
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                // Resetează UI în caz de eroare
                btnVizualizeaza.Text = "Vizualizeaza in Excel";
                btnVizualizeaza.Enabled = true;
                this.Cursor = Cursors.Default;

                MessageBox.Show($"Eroare la deschiderea raportului: {ex.Message}\n\nDetalii: {ex.GetType().Name}", "Eroare", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CurataFormular()
        {
            txtNume.Clear();
            txtDescriere.Clear();
            txtPret.Clear();
            txtStoc.Clear();
            txtNume.Focus();
        }
    }
}
