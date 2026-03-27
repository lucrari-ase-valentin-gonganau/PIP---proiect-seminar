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
using ProiectIngineriaProgramarii.Models;
using ProiectIngineriaProgramarii.WordAddin;

namespace ProiectIngineriaProgramarii
{
    public partial class FacturiForm : Form
    {
        private StartForm _mainForm;
        private DatabaseManager _dbManager;
        private ClientRepository _clientRepository;
        private ProdusRepository _produsRepository;
        private FacturaRepository _facturaRepository;
        private FacturaWordGenerator _wordGenerator;

        private Factura _facturaActuala;
        private BindingList<ItemFactura> _itemuriFactura;

        public FacturiForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Gestiune Facturi";
            _mainForm = mainForm;
            InitializeRepositories();
            InitializeFactura();
            IncarcaDate();
        }

        public FacturiForm()
        {
            InitializeComponent();
            this.Text = "Gestiune Facturi";
            InitializeRepositories();
            InitializeFactura();
            IncarcaDate();
        }

        private void InitializeRepositories()
        {
            _dbManager = new DatabaseManager();
            _clientRepository = new ClientRepository(_dbManager);
            _produsRepository = new ProdusRepository(_dbManager);
            _facturaRepository = new FacturaRepository(_dbManager, _clientRepository);
            _wordGenerator = new FacturaWordGenerator();
        }

        private void InitializeFactura()
        {
            _facturaActuala = new Factura
            {
                NumarFactura = _facturaRepository.GenerareNumarFactura(),
                DataEmitere = DateTime.Now,
                Status = "Emisa"
            };

            _itemuriFactura = new BindingList<ItemFactura>();
            dgvItemuri.DataSource = _itemuriFactura;

            ConfigureazaDataGridView();
        }

        private void ConfigureazaDataGridView()
        {
            if (dgvItemuri.Columns.Count > 0)
            {
                dgvItemuri.Columns["Id"].Visible = false;
                dgvItemuri.Columns["FacturaId"].Visible = false;
                dgvItemuri.Columns["ProdusId"].Visible = false;
                dgvItemuri.Columns["NumeProdus"].HeaderText = "Produs";
                dgvItemuri.Columns["Cantitate"].HeaderText = "Cantitate";
                dgvItemuri.Columns["PretUnitar"].HeaderText = "Pret Unitar";
                dgvItemuri.Columns["Subtotal"].HeaderText = "Subtotal";

                dgvItemuri.Columns["PretUnitar"].DefaultCellStyle.Format = "N2";
                dgvItemuri.Columns["Subtotal"].DefaultCellStyle.Format = "N2";
            }
        }

        private void IncarcaDate()
        {
            try
            {
                var clienti = _clientRepository.GetAll();
                cmbClient.DataSource = clienti;
                cmbClient.DisplayMember = "NumeComplet";
                cmbClient.ValueMember = "Id";

                var produse = _produsRepository.GetAll();
                cmbProdus.DataSource = produse;
                cmbProdus.DisplayMember = "Nume";
                cmbProdus.ValueMember = "Id";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la incarcarea datelor: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbProdus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbProdus.SelectedItem is Produs produs)
            {
                if (produs.StocDisponibil <= 0)
                {
                    MessageBox.Show("Produsul selectat nu mai este disponibil pe stoc!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btnAdaugaProdus_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbProdus.SelectedItem == null)
                {
                    MessageBox.Show("Selecteaza un produs!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!int.TryParse(txtCantitate.Text, out int cantitate) || cantitate <= 0)
                {
                    MessageBox.Show("Cantitatea trebuie sa fie un numar mai mare decat 0!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var produs = (Produs)cmbProdus.SelectedItem;

                if (cantitate > produs.StocDisponibil)
                {
                    MessageBox.Show($"Stoc insuficient! Disponibil: {produs.StocDisponibil}", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var itemExistent = _itemuriFactura.FirstOrDefault(i => i.ProdusId == produs.Id);
                if (itemExistent != null)
                {
                    itemExistent.Cantitate += cantitate;
                    itemExistent.CalculeazaSubtotal();
                }
                else
                {
                    var itemFactura = new ItemFactura
                    {
                        ProdusId = produs.Id,
                        NumeProdus = produs.Nume,
                        Cantitate = cantitate,
                        PretUnitar = produs.Pret
                    };
                    itemFactura.CalculeazaSubtotal();
                    _itemuriFactura.Add(itemFactura);
                }

                ConfigureazaDataGridView();
                txtCantitate.Text = "1";
                CalculeazaTotaluri();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la adaugarea produsului: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStergeProdus_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvItemuri.SelectedRows.Count > 0)
                {
                    var index = dgvItemuri.SelectedRows[0].Index;
                    _itemuriFactura.RemoveAt(index);
                    CalculeazaTotaluri();
                }
                else
                {
                    MessageBox.Show("Selecteaza un produs pentru a-l sterge!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la stergerea produsului: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CalculeazaTotaluri()
        {
            _facturaActuala.Itemi = _itemuriFactura.ToList();
            _facturaActuala.CalculeazaTotaluri();

            lblSubtotalValue.Text = $"{_facturaActuala.Subtotal:N2} RON";
            lblTVAValue.Text = $"{_facturaActuala.TVA:N2} RON";
            lblTotalValue.Text = $"{_facturaActuala.Total:N2} RON";
        }

        private void btnSalveazaFactura_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbClient.SelectedItem == null)
                {
                    MessageBox.Show("Selecteaza un client!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (_itemuriFactura.Count == 0)
                {
                    MessageBox.Show("Adauga cel putin un produs!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _facturaActuala.ClientId = (int)cmbClient.SelectedValue;
                _facturaActuala.Client = (Client)cmbClient.SelectedItem;
                _facturaActuala.Observatii = txtObservatii.Text;
                _facturaActuala.Itemi = _itemuriFactura.ToList();

                _facturaRepository.Add(_facturaActuala);

                MessageBox.Show($"Factura {_facturaActuala.NumarFactura} a fost salvata cu succes!", "Succes",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                InitializeFactura();
                txtObservatii.Clear();
                CalculeazaTotaluri();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la salvarea facturii: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenereazaWord_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbClient.SelectedItem == null)
                {
                    MessageBox.Show("Selecteaza un client!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (_itemuriFactura.Count == 0)
                {
                    MessageBox.Show("Adauga cel putin un produs!", "Atentie",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _facturaActuala.ClientId = (int)cmbClient.SelectedValue;
                _facturaActuala.Client = (Client)cmbClient.SelectedItem;
                _facturaActuala.Observatii = txtObservatii.Text;
                _facturaActuala.Itemi = _itemuriFactura.ToList();

                this.Cursor = Cursors.WaitCursor;
                btnGenereazaWord.Enabled = false;
                btnGenereazaWord.Text = "Se genereaza Word...";
                Application.DoEvents();

                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string fileName = $"Factura_{_facturaActuala.NumarFactura}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                string filePath = Path.Combine(documentsPath, fileName);

                _wordGenerator.GenereazaFactura(_facturaActuala, filePath);

                btnGenereazaWord.Text = "Genereaza Word";
                btnGenereazaWord.Enabled = true;
                this.Cursor = Cursors.Default;

                var result = MessageBox.Show(
                    $"Factura a fost generata cu succes!\n\nLocatie: {filePath}\n\nDoriti sa deschideti factura?",
                    "Succes",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    _wordGenerator.DeschideFactura(filePath);
                }
            }
            catch (Exception ex)
            {
                btnGenereazaWord.Text = "Genereaza Word";
                btnGenereazaWord.Enabled = true;
                this.Cursor = Cursors.Default;

                MessageBox.Show($"Eroare la generarea facturii: {ex.Message}", "Eroare",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
