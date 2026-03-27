namespace ProiectIngineriaProgramarii
{
    partial class ClientiForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            lblTitle = new Label();
            lblNume = new Label();
            txtNume = new TextBox();
            lblPrenume = new Label();
            txtPrenume = new TextBox();
            lblEmail = new Label();
            txtEmail = new TextBox();
            lblTelefon = new Label();
            txtTelefon = new TextBox();
            lblAdresa = new Label();
            txtAdresa = new TextBox();
            btnAdauga = new Button();
            btnVizualizeaza = new Button();
            dgvClienti = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)dgvClienti).BeginInit();
            SuspendLayout();
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Yu Gothic", 18F, FontStyle.Bold);
            lblTitle.Location = new Point(130, 20);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(202, 31);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "Gestiune Clienți";
            // 
            // lblNume
            // 
            lblNume.AutoSize = true;
            lblNume.Font = new Font("Yu Gothic", 10F);
            lblNume.Location = new Point(130, 80);
            lblNume.Name = "lblNume";
            lblNume.Size = new Size(50, 18);
            lblNume.TabIndex = 1;
            lblNume.Text = "Nume:";
            // 
            // txtNume
            // 
            txtNume.Font = new Font("Yu Gothic", 10F);
            txtNume.Location = new Point(220, 77);
            txtNume.Name = "txtNume";
            txtNume.Size = new Size(250, 29);
            txtNume.TabIndex = 2;
            // 
            // lblPrenume
            // 
            lblPrenume.AutoSize = true;
            lblPrenume.Font = new Font("Yu Gothic", 10F);
            lblPrenume.Location = new Point(130, 120);
            lblPrenume.Name = "lblPrenume";
            lblPrenume.Size = new Size(70, 18);
            lblPrenume.TabIndex = 3;
            lblPrenume.Text = "Prenume:";
            // 
            // txtPrenume
            // 
            txtPrenume.Font = new Font("Yu Gothic", 10F);
            txtPrenume.Location = new Point(220, 117);
            txtPrenume.Name = "txtPrenume";
            txtPrenume.Size = new Size(250, 29);
            txtPrenume.TabIndex = 4;
            // 
            // lblEmail
            // 
            lblEmail.AutoSize = true;
            lblEmail.Font = new Font("Yu Gothic", 10F);
            lblEmail.Location = new Point(130, 160);
            lblEmail.Name = "lblEmail";
            lblEmail.Size = new Size(49, 18);
            lblEmail.TabIndex = 5;
            lblEmail.Text = "Email:";
            // 
            // txtEmail
            // 
            txtEmail.Font = new Font("Yu Gothic", 10F);
            txtEmail.Location = new Point(220, 157);
            txtEmail.Name = "txtEmail";
            txtEmail.Size = new Size(250, 29);
            txtEmail.TabIndex = 6;
            // 
            // lblTelefon
            // 
            lblTelefon.AutoSize = true;
            lblTelefon.Font = new Font("Yu Gothic", 10F);
            lblTelefon.Location = new Point(130, 200);
            lblTelefon.Name = "lblTelefon";
            lblTelefon.Size = new Size(61, 18);
            lblTelefon.TabIndex = 7;
            lblTelefon.Text = "Telefon:";
            // 
            // txtTelefon
            // 
            txtTelefon.Font = new Font("Yu Gothic", 10F);
            txtTelefon.Location = new Point(220, 197);
            txtTelefon.Name = "txtTelefon";
            txtTelefon.Size = new Size(250, 29);
            txtTelefon.TabIndex = 8;
            // 
            // lblAdresa
            // 
            lblAdresa.AutoSize = true;
            lblAdresa.Font = new Font("Yu Gothic", 10F);
            lblAdresa.Location = new Point(130, 240);
            lblAdresa.Name = "lblAdresa";
            lblAdresa.Size = new Size(57, 18);
            lblAdresa.TabIndex = 9;
            lblAdresa.Text = "Adresa:";
            // 
            // txtAdresa
            // 
            txtAdresa.Font = new Font("Yu Gothic", 10F);
            txtAdresa.Location = new Point(220, 237);
            txtAdresa.Multiline = true;
            txtAdresa.Name = "txtAdresa";
            txtAdresa.Size = new Size(250, 60);
            txtAdresa.TabIndex = 10;
            // 
            // btnAdauga
            // 
            btnAdauga.BackColor = Color.FromArgb(76, 175, 80);
            btnAdauga.FlatStyle = FlatStyle.Flat;
            btnAdauga.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            btnAdauga.ForeColor = Color.White;
            btnAdauga.Location = new Point(220, 316);
            btnAdauga.Name = "btnAdauga";
            btnAdauga.Size = new Size(250, 40);
            btnAdauga.TabIndex = 11;
            btnAdauga.Text = "Adaugă Client";
            btnAdauga.UseVisualStyleBackColor = false;
            btnAdauga.Click += btnAdauga_Click;
            // 
            // btnVizualizeaza
            // 
            btnVizualizeaza.BackColor = Color.FromArgb(33, 150, 243);
            btnVizualizeaza.FlatStyle = FlatStyle.Flat;
            btnVizualizeaza.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            btnVizualizeaza.ForeColor = Color.White;
            btnVizualizeaza.Location = new Point(220, 362);
            btnVizualizeaza.Name = "btnVizualizeaza";
            btnVizualizeaza.Size = new Size(250, 40);
            btnVizualizeaza.TabIndex = 12;
            btnVizualizeaza.Text = "Vizualizeaza in Excel";
            btnVizualizeaza.UseVisualStyleBackColor = false;
            btnVizualizeaza.Click += btnVizualizeaza_Click;
            // 
            // dgvClienti
            // 
            dgvClienti.AllowUserToAddRows = false;
            dgvClienti.AllowUserToDeleteRows = false;
            dgvClienti.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvClienti.BackgroundColor = Color.White;
            dgvClienti.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvClienti.Location = new Point(500, 77);
            dgvClienti.Name = "dgvClienti";
            dgvClienti.ReadOnly = true;
            dgvClienti.RowHeadersWidth = 51;
            dgvClienti.Size = new Size(450, 350);
            dgvClienti.TabIndex = 13;
            // 
            // ClientiForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(984, 461);
            Controls.Add(dgvClienti);
            Controls.Add(btnVizualizeaza);
            Controls.Add(btnAdauga);
            Controls.Add(txtAdresa);
            Controls.Add(lblAdresa);
            Controls.Add(txtTelefon);
            Controls.Add(lblTelefon);
            Controls.Add(txtEmail);
            Controls.Add(lblEmail);
            Controls.Add(txtPrenume);
            Controls.Add(lblPrenume);
            Controls.Add(txtNume);
            Controls.Add(lblNume);
            Controls.Add(lblTitle);
            Name = "ClientiForm";
            Text = "Gestiune Clienți";
            ((System.ComponentModel.ISupportInitialize)dgvClienti).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblTitle;
        private Label lblNume;
        private TextBox txtNume;
        private Label lblPrenume;
        private TextBox txtPrenume;
        private Label lblEmail;
        private TextBox txtEmail;
        private Label lblTelefon;
        private TextBox txtTelefon;
        private Label lblAdresa;
        private TextBox txtAdresa;
        private Button btnAdauga;
        private Button btnVizualizeaza;
        private DataGridView dgvClienti;
    }
}
