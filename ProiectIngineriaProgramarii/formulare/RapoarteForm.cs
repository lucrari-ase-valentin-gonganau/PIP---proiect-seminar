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

        private void btnVeziGraficExcel_Click(object sender, EventArgs e)
        {
            try
            {
                var facturi = _facturaRepository.GetAll();

                if (facturi == null || facturi.Count == 0)
                {
                    MessageBox.Show("Nu există facturi în baza de date pentru a genera raportul.",
                        "Informație", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Raport_Vanzari_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                    Title = "Salvează raportul Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var generator = new RapoarteExcelGenerator();
                    generator.GenereazaRaportComplet(facturi, saveFileDialog.FileName);

                    var result = MessageBox.Show(
                        "Raportul Excel a fost generat cu succes!\n\nDoriți să deschideți fișierul?",
                        "Succes",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = saveFileDialog.FileName,
                            UseShellExecute = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Eroare la generarea raportului: {ex.Message}",
                    "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
