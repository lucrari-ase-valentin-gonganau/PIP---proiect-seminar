namespace ProiectIngineriaProgramarii
{
    partial class FacturiForm
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
            lblClient = new Label();
            cmbClient = new ComboBox();
            groupBoxProduse = new GroupBox();
            btnAdaugaProdus = new Button();
            txtCantitate = new TextBox();
            lblCantitate = new Label();
            cmbProdus = new ComboBox();
            lblProdusSelect = new Label();
            dgvItemuri = new DataGridView();
            btnStergeProdus = new Button();
            groupBoxTotalizare = new GroupBox();
            lblTotalValue = new Label();
            lblTotal = new Label();
            lblTVAValue = new Label();
            lblTVA = new Label();
            lblSubtotalValue = new Label();
            lblSubtotal = new Label();
            lblObservatii = new Label();
            txtObservatii = new TextBox();
            btnSalveazaFactura = new Button();
            btnGenereazaWord = new Button();
            groupBoxProduse.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvItemuri).BeginInit();
            groupBoxTotalizare.SuspendLayout();
            SuspendLayout();
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Yu Gothic", 18F, FontStyle.Bold);
            lblTitle.Location = new Point(30, 20);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(255, 31);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "Creare Factura Noua";
            // 
            // lblClient
            // 
            lblClient.AutoSize = true;
            lblClient.Font = new Font("Yu Gothic", 10F);
            lblClient.Location = new Point(30, 70);
            lblClient.Name = "lblClient";
            lblClient.Size = new Size(126, 18);
            lblClient.TabIndex = 1;
            lblClient.Text = "Selecteaza Client:";
            // 
            // cmbClient
            // 
            cmbClient.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbClient.Font = new Font("Yu Gothic", 10F);
            cmbClient.FormattingEnabled = true;
            cmbClient.Location = new Point(30, 95);
            cmbClient.Name = "cmbClient";
            cmbClient.Size = new Size(300, 25);
            cmbClient.TabIndex = 2;
            // 
            // groupBoxProduse
            // 
            groupBoxProduse.Controls.Add(btnAdaugaProdus);
            groupBoxProduse.Controls.Add(txtCantitate);
            groupBoxProduse.Controls.Add(lblCantitate);
            groupBoxProduse.Controls.Add(cmbProdus);
            groupBoxProduse.Controls.Add(lblProdusSelect);
            groupBoxProduse.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            groupBoxProduse.Location = new Point(30, 140);
            groupBoxProduse.Name = "groupBoxProduse";
            groupBoxProduse.Size = new Size(470, 130);
            groupBoxProduse.TabIndex = 3;
            groupBoxProduse.TabStop = false;
            groupBoxProduse.Text = "Adauga Produse";
            // 
            // btnAdaugaProdus
            // 
            btnAdaugaProdus.BackColor = Color.FromArgb(76, 175, 80);
            btnAdaugaProdus.FlatStyle = FlatStyle.Flat;
            btnAdaugaProdus.Font = new Font("Yu Gothic", 9F, FontStyle.Bold);
            btnAdaugaProdus.ForeColor = Color.White;
            btnAdaugaProdus.Location = new Point(370, 45);
            btnAdaugaProdus.Name = "btnAdaugaProdus";
            btnAdaugaProdus.Size = new Size(85, 30);
            btnAdaugaProdus.TabIndex = 4;
            btnAdaugaProdus.Text = "Adauga";
            btnAdaugaProdus.UseVisualStyleBackColor = false;
            btnAdaugaProdus.Click += btnAdaugaProdus_Click;
            // 
            // txtCantitate
            // 
            txtCantitate.Font = new Font("Yu Gothic", 9F);
            txtCantitate.Location = new Point(280, 50);
            txtCantitate.Name = "txtCantitate";
            txtCantitate.Size = new Size(80, 27);
            txtCantitate.TabIndex = 3;
            txtCantitate.Text = "1";
            // 
            // lblCantitate
            // 
            lblCantitate.AutoSize = true;
            lblCantitate.Font = new Font("Yu Gothic", 9F);
            lblCantitate.Location = new Point(280, 30);
            lblCantitate.Name = "lblCantitate";
            lblCantitate.Size = new Size(61, 16);
            lblCantitate.TabIndex = 2;
            lblCantitate.Text = "Cantitate:";
            // 
            // cmbProdus
            // 
            cmbProdus.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbProdus.Font = new Font("Yu Gothic", 9F);
            cmbProdus.FormattingEnabled = true;
            cmbProdus.Location = new Point(15, 50);
            cmbProdus.Name = "cmbProdus";
            cmbProdus.Size = new Size(250, 24);
            cmbProdus.TabIndex = 1;
            cmbProdus.SelectedIndexChanged += cmbProdus_SelectedIndexChanged;
            // 
            // lblProdusSelect
            // 
            lblProdusSelect.AutoSize = true;
            lblProdusSelect.Font = new Font("Yu Gothic", 9F);
            lblProdusSelect.Location = new Point(15, 30);
            lblProdusSelect.Name = "lblProdusSelect";
            lblProdusSelect.Size = new Size(49, 16);
            lblProdusSelect.TabIndex = 0;
            lblProdusSelect.Text = "Produs:";
            // 
            // dgvItemuri
            // 
            dgvItemuri.AllowUserToAddRows = false;
            dgvItemuri.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvItemuri.BackgroundColor = Color.White;
            dgvItemuri.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvItemuri.Location = new Point(30, 285);
            dgvItemuri.Name = "dgvItemuri";
            dgvItemuri.ReadOnly = true;
            dgvItemuri.RowHeadersWidth = 51;
            dgvItemuri.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvItemuri.Size = new Size(470, 200);
            dgvItemuri.TabIndex = 4;
            // 
            // btnStergeProdus
            // 
            btnStergeProdus.BackColor = Color.FromArgb(244, 67, 54);
            btnStergeProdus.FlatStyle = FlatStyle.Flat;
            btnStergeProdus.Font = new Font("Yu Gothic", 9F, FontStyle.Bold);
            btnStergeProdus.ForeColor = Color.White;
            btnStergeProdus.Location = new Point(30, 495);
            btnStergeProdus.Name = "btnStergeProdus";
            btnStergeProdus.Size = new Size(120, 35);
            btnStergeProdus.TabIndex = 5;
            btnStergeProdus.Text = "Sterge Produs";
            btnStergeProdus.UseVisualStyleBackColor = false;
            btnStergeProdus.Click += btnStergeProdus_Click;
            // 
            // groupBoxTotalizare
            // 
            groupBoxTotalizare.Controls.Add(lblTotalValue);
            groupBoxTotalizare.Controls.Add(lblTotal);
            groupBoxTotalizare.Controls.Add(lblTVAValue);
            groupBoxTotalizare.Controls.Add(lblTVA);
            groupBoxTotalizare.Controls.Add(lblSubtotalValue);
            groupBoxTotalizare.Controls.Add(lblSubtotal);
            groupBoxTotalizare.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            groupBoxTotalizare.Location = new Point(520, 70);
            groupBoxTotalizare.Name = "groupBoxTotalizare";
            groupBoxTotalizare.Size = new Size(300, 150);
            groupBoxTotalizare.TabIndex = 6;
            groupBoxTotalizare.TabStop = false;
            groupBoxTotalizare.Text = "Totalizare";
            // 
            // lblTotalValue
            // 
            lblTotalValue.AutoSize = true;
            lblTotalValue.Font = new Font("Yu Gothic", 12F, FontStyle.Bold);
            lblTotalValue.ForeColor = Color.FromArgb(76, 175, 80);
            lblTotalValue.Location = new Point(155, 100);
            lblTotalValue.Name = "lblTotalValue";
            lblTotalValue.Size = new Size(82, 21);
            lblTotalValue.TabIndex = 5;
            lblTotalValue.Text = "0.00 RON";
            // 
            // lblTotal
            // 
            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Yu Gothic", 12F, FontStyle.Bold);
            lblTotal.Location = new Point(20, 100);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(69, 21);
            lblTotal.TabIndex = 4;
            lblTotal.Text = "TOTAL:";
            // 
            // lblTVAValue
            // 
            lblTVAValue.AutoSize = true;
            lblTVAValue.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            lblTVAValue.Location = new Point(180, 65);
            lblTVAValue.Name = "lblTVAValue";
            lblTVAValue.Size = new Size(72, 18);
            lblTVAValue.TabIndex = 3;
            lblTVAValue.Text = "0.00 RON";
            // 
            // lblTVA
            // 
            lblTVA.AutoSize = true;
            lblTVA.Font = new Font("Yu Gothic", 10F);
            lblTVA.Location = new Point(20, 65);
            lblTVA.Name = "lblTVA";
            lblTVA.Size = new Size(79, 18);
            lblTVA.TabIndex = 2;
            lblTVA.Text = "TVA (19%):";
            // 
            // lblSubtotalValue
            // 
            lblSubtotalValue.AutoSize = true;
            lblSubtotalValue.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            lblSubtotalValue.Location = new Point(180, 35);
            lblSubtotalValue.Name = "lblSubtotalValue";
            lblSubtotalValue.Size = new Size(72, 18);
            lblSubtotalValue.TabIndex = 1;
            lblSubtotalValue.Text = "0.00 RON";
            // 
            // lblSubtotal
            // 
            lblSubtotal.AutoSize = true;
            lblSubtotal.Font = new Font("Yu Gothic", 10F);
            lblSubtotal.Location = new Point(20, 35);
            lblSubtotal.Name = "lblSubtotal";
            lblSubtotal.Size = new Size(67, 18);
            lblSubtotal.TabIndex = 0;
            lblSubtotal.Text = "Subtotal:";
            // 
            // lblObservatii
            // 
            lblObservatii.AutoSize = true;
            lblObservatii.Font = new Font("Yu Gothic", 10F);
            lblObservatii.Location = new Point(520, 240);
            lblObservatii.Name = "lblObservatii";
            lblObservatii.Size = new Size(78, 18);
            lblObservatii.TabIndex = 7;
            lblObservatii.Text = "Observatii:";
            // 
            // txtObservatii
            // 
            txtObservatii.Font = new Font("Yu Gothic", 9F);
            txtObservatii.Location = new Point(520, 265);
            txtObservatii.Multiline = true;
            txtObservatii.Name = "txtObservatii";
            txtObservatii.Size = new Size(300, 100);
            txtObservatii.TabIndex = 8;
            // 
            // btnSalveazaFactura
            // 
            btnSalveazaFactura.BackColor = Color.FromArgb(33, 150, 243);
            btnSalveazaFactura.FlatStyle = FlatStyle.Flat;
            btnSalveazaFactura.Font = new Font("Yu Gothic", 11F, FontStyle.Bold);
            btnSalveazaFactura.ForeColor = Color.White;
            btnSalveazaFactura.Location = new Point(520, 385);
            btnSalveazaFactura.Name = "btnSalveazaFactura";
            btnSalveazaFactura.Size = new Size(300, 45);
            btnSalveazaFactura.TabIndex = 9;
            btnSalveazaFactura.Text = "Salveaza Factura";
            btnSalveazaFactura.UseVisualStyleBackColor = false;
            btnSalveazaFactura.Click += btnSalveazaFactura_Click;
            // 
            // btnGenereazaWord
            // 
            btnGenereazaWord.BackColor = Color.FromArgb(255, 152, 0);
            btnGenereazaWord.FlatStyle = FlatStyle.Flat;
            btnGenereazaWord.Font = new Font("Yu Gothic", 11F, FontStyle.Bold);
            btnGenereazaWord.ForeColor = Color.White;
            btnGenereazaWord.Location = new Point(520, 440);
            btnGenereazaWord.Name = "btnGenereazaWord";
            btnGenereazaWord.Size = new Size(300, 45);
            btnGenereazaWord.TabIndex = 10;
            btnGenereazaWord.Text = "Genereaza factura in Word";
            btnGenereazaWord.UseVisualStyleBackColor = false;
            btnGenereazaWord.Click += btnGenereazaWord_Click;
            // 
            // FacturiForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(850, 560);
            Controls.Add(btnGenereazaWord);
            Controls.Add(btnSalveazaFactura);
            Controls.Add(txtObservatii);
            Controls.Add(lblObservatii);
            Controls.Add(groupBoxTotalizare);
            Controls.Add(btnStergeProdus);
            Controls.Add(dgvItemuri);
            Controls.Add(groupBoxProduse);
            Controls.Add(cmbClient);
            Controls.Add(lblClient);
            Controls.Add(lblTitle);
            Name = "FacturiForm";
            Text = "Gestiune Facturi";
            groupBoxProduse.ResumeLayout(false);
            groupBoxProduse.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvItemuri).EndInit();
            groupBoxTotalizare.ResumeLayout(false);
            groupBoxTotalizare.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblTitle;
        private Label lblClient;
        private ComboBox cmbClient;
        private GroupBox groupBoxProduse;
        private Button btnAdaugaProdus;
        private TextBox txtCantitate;
        private Label lblCantitate;
        private ComboBox cmbProdus;
        private Label lblProdusSelect;
        private DataGridView dgvItemuri;
        private Button btnStergeProdus;
        private GroupBox groupBoxTotalizare;
        private Label lblTotalValue;
        private Label lblTotal;
        private Label lblTVAValue;
        private Label lblTVA;
        private Label lblSubtotalValue;
        private Label lblSubtotal;
        private Label lblObservatii;
        private TextBox txtObservatii;
        private Button btnSalveazaFactura;
        private Button btnGenereazaWord;
    }
}
