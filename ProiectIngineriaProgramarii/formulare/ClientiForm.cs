using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ProiectIngineriaProgramarii.Data;
using ProiectIngineriaProgramarii.Models;
using ProiectIngineriaProgramarii.ExcelAddin;

namespace ProiectIngineriaProgramarii
{
    public partial class ClientiForm : Form
    {
        private StartForm _mainForm;
        private DatabaseManager _dbManager;
        private ClientRepository _clientRepository;
        private RapoarteExcelGenerator _excelGenerator;

        public ClientiForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Gestiune Clienti";
            _mainForm = mainForm;
            _dbManager = new DatabaseManager();
            _clientRepository = new ClientRepository(_dbManager);
            _excelGenerator = new RapoarteExcelGenerator();

            IncarcaClienti();
        }

        public ClientiForm()
        {
            InitializeComponent();
            this.Text = "Gestiune Clienti";
            _dbManager = new DatabaseManager();
            _clientRepository = new ClientRepository(_dbManager);
            _excelGenerator = new RapoarteExcelGenerator();
            IncarcaClienti();
        }

        private void IncarcaClienti()
        {
            try
            {
                var clienti = _clientRepository.GetAll();
                dgvClienti.DataSource = clienti;

                if (dgvClienti.Columns.Count > 0)
                {
                    dgvClienti.Columns["Id"].HeaderText = "ID";
                    dgvClienti.Columns["Nume"].HeaderText = "Nume";
                    dgvClienti.Columns["Prenume"].HeaderText = "Prenume";
                    dgvClienti.Columns["Email"].HeaderText = "Email";
                    dgvClienti.Columns["Telefon"].HeaderText = "Telefon";
                    dgvClienti.Columns["Adresa"].HeaderText = "Adresa";
                    dgvClienti.Columns["DataInregistrare"].HeaderText = "Data Inregistrare";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la incarcarea clientilor: {ex.Message}", "Eroare", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdauga_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtNume.Text) || string.IsNullOrWhiteSpace(txtPrenume.Text))
                {
                    MessageBox.Show("Numele si prenumele sunt obligatorii!", "Atentie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var client = new Client
                {
                    Nume = txtNume.Text.Trim(),
                    Prenume = txtPrenume.Text.Trim(),
                    Email = txtEmail.Text.Trim(),
                    Telefon = txtTelefon.Text.Trim(),
                    Adresa = txtAdresa.Text.Trim(),
                    DataInregistrare = DateTime.Now
                };

                _clientRepository.Add(client);
                MessageBox.Show("Client adaugat cu succes!", "Succes", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                CurataFormular();
                IncarcaClienti();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la adaugarea clientului: {ex.Message}", "Eroare", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVizualizeaza_Click(object sender, EventArgs e)
        {
            try
            {
                var clienti = _clientRepository.GetAll();

                if (clienti.Count == 0)
                {
                    MessageBox.Show("Nu exista clienti in baza de date!", "Informatie", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Afișează loader
                this.Cursor = Cursors.WaitCursor;
                btnVizualizeaza.Enabled = false;
                btnVizualizeaza.Text = "Se deschide Excel...";
                Application.DoEvents();

                // Deschide Excel
                _excelGenerator.DeschideRaportClientiDirect(clienti);

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
            txtPrenume.Clear();
            txtEmail.Clear();
            txtTelefon.Clear();
            txtAdresa.Clear();
            txtNume.Focus();
        }
    }
}
