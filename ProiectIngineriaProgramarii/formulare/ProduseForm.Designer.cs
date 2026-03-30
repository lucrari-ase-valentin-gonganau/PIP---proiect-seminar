namespace ProiectIngineriaProgramarii
{
    partial class ProduseForm
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
            lblDescriere = new Label();
            txtDescriere = new TextBox();
            lblPret = new Label();
            txtPret = new TextBox();
            lblStoc = new Label();
            txtStoc = new TextBox();
            btnAdauga = new Button();
            btnVizualizeaza = new Button();
            dgvProduse = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)dgvProduse).BeginInit();
            SuspendLayout();
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Yu Gothic", 18F, FontStyle.Bold);
            lblTitle.Location = new Point(130, 20);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(257, 31);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "Gestiune de produse";
            // 
            // lblNume
            // 
            lblNume.AutoSize = true;
            lblNume.Font = new Font("Yu Gothic", 10F);
            lblNume.Location = new Point(130, 80);
            lblNume.Name = "lblNume";
            lblNume.Size = new Size(99, 18);
            lblNume.TabIndex = 1;
            lblNume.Text = "Nume Produs:";
            // 
            // txtNume
            // 
            txtNume.Font = new Font("Yu Gothic", 10F);
            txtNume.Location = new Point(240, 77);
            txtNume.Name = "txtNume";
            txtNume.Size = new Size(230, 29);
            txtNume.TabIndex = 2;
            // 
            // lblDescriere
            // 
            lblDescriere.AutoSize = true;
            lblDescriere.Font = new Font("Yu Gothic", 10F);
            lblDescriere.Location = new Point(130, 120);
            lblDescriere.Name = "lblDescriere";
            lblDescriere.Size = new Size(74, 18);
            lblDescriere.TabIndex = 3;
            lblDescriere.Text = "Descriere:";
            // 
            // txtDescriere
            // 
            txtDescriere.Font = new Font("Yu Gothic", 10F);
            txtDescriere.Location = new Point(240, 117);
            txtDescriere.Multiline = true;
            txtDescriere.Name = "txtDescriere";
            txtDescriere.Size = new Size(230, 60);
            txtDescriere.TabIndex = 4;
            // 
            // lblPret
            // 
            lblPret.AutoSize = true;
            lblPret.Font = new Font("Yu Gothic", 10F);
            lblPret.Location = new Point(130, 195);
            lblPret.Name = "lblPret";
            lblPret.Size = new Size(82, 18);
            lblPret.TabIndex = 5;
            lblPret.Text = "Pret (RON):";
            // 
            // txtPret
            // 
            txtPret.Font = new Font("Yu Gothic", 10F);
            txtPret.Location = new Point(240, 192);
            txtPret.Name = "txtPret";
            txtPret.Size = new Size(230, 29);
            txtPret.TabIndex = 6;
            // 
            // lblStoc
            // 
            lblStoc.AutoSize = true;
            lblStoc.Font = new Font("Yu Gothic", 10F);
            lblStoc.Location = new Point(130, 235);
            lblStoc.Name = "lblStoc";
            lblStoc.Size = new Size(110, 18);
            lblStoc.TabIndex = 7;
            lblStoc.Text = "Stoc Disponibil:";
            // 
            // txtStoc
            // 
            txtStoc.Font = new Font("Yu Gothic", 10F);
            txtStoc.Location = new Point(240, 232);
            txtStoc.Name = "txtStoc";
            txtStoc.Size = new Size(230, 29);
            txtStoc.TabIndex = 8;
            // 
            // btnAdauga
            // 
            btnAdauga.BackColor = Color.FromArgb(76, 175, 80);
            btnAdauga.FlatStyle = FlatStyle.Flat;
            btnAdauga.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            btnAdauga.ForeColor = Color.White;
            btnAdauga.Location = new Point(240, 280);
            btnAdauga.Name = "btnAdauga";
            btnAdauga.Size = new Size(230, 40);
            btnAdauga.TabIndex = 9;
            btnAdauga.Text = "Adauga un produs";
            btnAdauga.UseVisualStyleBackColor = false;
            btnAdauga.Click += btnAdauga_Click;
            // 
            // btnVizualizeaza
            // 
            btnVizualizeaza.BackColor = Color.FromArgb(33, 150, 243);
            btnVizualizeaza.FlatStyle = FlatStyle.Flat;
            btnVizualizeaza.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            btnVizualizeaza.ForeColor = Color.White;
            btnVizualizeaza.Location = new Point(240, 339);
            btnVizualizeaza.Name = "btnVizualizeaza";
            btnVizualizeaza.Size = new Size(230, 40);
            btnVizualizeaza.TabIndex = 10;
            btnVizualizeaza.Text = "Vizualizeaza in Excel";
            btnVizualizeaza.UseVisualStyleBackColor = false;
            btnVizualizeaza.Click += btnVizualizeaza_Click;
            // 
            // dgvProduse
            // 
            dgvProduse.AllowUserToAddRows = false;
            dgvProduse.AllowUserToDeleteRows = false;
            dgvProduse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvProduse.BackgroundColor = Color.White;
            dgvProduse.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvProduse.Location = new Point(500, 77);
            dgvProduse.Name = "dgvProduse";
            dgvProduse.ReadOnly = true;
            dgvProduse.RowHeadersWidth = 51;
            dgvProduse.Size = new Size(450, 350);
            dgvProduse.TabIndex = 11;
            // 
            // ProduseForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(984, 461);
            Controls.Add(dgvProduse);
            Controls.Add(btnVizualizeaza);
            Controls.Add(btnAdauga);
            Controls.Add(txtStoc);
            Controls.Add(lblStoc);
            Controls.Add(txtPret);
            Controls.Add(lblPret);
            Controls.Add(txtDescriere);
            Controls.Add(lblDescriere);
            Controls.Add(txtNume);
            Controls.Add(lblNume);
            Controls.Add(lblTitle);
            Name = "ProduseForm";
            Text = "Gestiune Produse";
            ((System.ComponentModel.ISupportInitialize)dgvProduse).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblTitle;
        private Label lblNume;
        private TextBox txtNume;
        private Label lblDescriere;
        private TextBox txtDescriere;
        private Label lblPret;
        private TextBox txtPret;
        private Label lblStoc;
        private TextBox txtStoc;
        private Button btnAdauga;
        private Button btnVizualizeaza;
        private DataGridView dgvProduse;
    }
}
