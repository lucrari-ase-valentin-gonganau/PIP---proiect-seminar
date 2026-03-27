namespace ProiectIngineriaProgramarii
{
    public partial class StartForm : Form
    {
        private Panel mainPanel;

        public StartForm()
        {
            InitializeComponent();
            CreateMainPanel();
            SetupMenuNavigation();
            ShowWelcomeScreen();
        }

        private void CreateMainPanel()
        {
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Location = new Point(0, menuStrip1.Height),
                Name = "mainPanel",
                TabIndex = 100,
                BackColor = Color.White
            };
            this.Controls.Add(mainPanel);
            mainPanel.SendToBack();
        }

        private void ShowWelcomeScreen()
        {
            lblWelcome.BringToFront();
            lblDescription.BringToFront();
        }

        private void SetupMenuNavigation()
        {
            clientiToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new ClientiForm(this));
            produseToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new ProduseForm(this));
            adaugaFacturaToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new FacturiForm(this));
            listaFacturiToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new ListaFacturiForm(this));
            rapoarteToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new RapoarteForm(this));
        }

        public void LoadFormInPanel(Form childForm)
        {
            lblWelcome.Visible = false;
            lblDescription.Visible = false;

            mainPanel.Controls.Clear();
            mainPanel.BringToFront();

            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;

            mainPanel.Controls.Add(childForm);
            mainPanel.Tag = childForm;
            childForm.Show();
        }

        public void ShowMainMenu()
        {
            mainPanel.Controls.Clear();
            mainPanel.SendToBack();
            lblWelcome.Visible = true;
            lblDescription.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
