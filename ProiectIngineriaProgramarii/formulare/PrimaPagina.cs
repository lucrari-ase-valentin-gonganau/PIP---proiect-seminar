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
        }

        private void CreateMainPanel()
        {
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Location = new Point(0, menuStrip1.Height),
                Name = "mainPanel",
                TabIndex = 100
            };
            this.Controls.Add(mainPanel);
            mainPanel.BringToFront();
        }

        private void SetupMenuNavigation()
        {
            clientiToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new ClientiForm(this));
            produseToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new ProduseForm(this));
            facturiToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new FacturiForm(this));
            rapoarteToolStripMenuItem.Click += (s, e) => LoadFormInPanel(new RapoarteForm(this));
        }

        public void LoadFormInPanel(Form childForm)
        {
            mainPanel.Controls.Clear();

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
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
