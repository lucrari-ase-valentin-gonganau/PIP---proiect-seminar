namespace ProiectIngineriaProgramarii
{
    partial class ListaFacturiForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dgvFacturi = new DataGridView();
            btnGenereazaWord = new Button();
            lblTitle = new Label();
            panelBottom = new Panel();
            ((System.ComponentModel.ISupportInitialize)dgvFacturi).BeginInit();
            panelBottom.SuspendLayout();
            SuspendLayout();
            // 
            // dgvFacturi
            // 
            dgvFacturi.AllowUserToAddRows = false;
            dgvFacturi.AllowUserToDeleteRows = false;
            dgvFacturi.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvFacturi.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvFacturi.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvFacturi.Location = new Point(20, 60);
            dgvFacturi.MultiSelect = false;
            dgvFacturi.Name = "dgvFacturi";
            dgvFacturi.ReadOnly = true;
            dgvFacturi.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvFacturi.Size = new Size(760, 320);
            dgvFacturi.TabIndex = 0;
            // 
            // btnGenereazaWord
            // 
            btnGenereazaWord.BackColor = Color.FromArgb(0, 122, 204);
            btnGenereazaWord.FlatStyle = FlatStyle.Flat;
            btnGenereazaWord.Font = new Font("Yu Gothic", 10F, FontStyle.Bold);
            btnGenereazaWord.ForeColor = Color.White;
            btnGenereazaWord.Location = new Point(20, 10);
            btnGenereazaWord.Name = "btnGenereazaWord";
            btnGenereazaWord.Size = new Size(336, 40);
            btnGenereazaWord.TabIndex = 1;
            btnGenereazaWord.Text = "Genereaza factura in Word";
            btnGenereazaWord.UseVisualStyleBackColor = false;
            btnGenereazaWord.Click += btnGenereazaWord_Click;
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Yu Gothic", 14F, FontStyle.Bold);
            lblTitle.Location = new Point(20, 20);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(125, 25);
            lblTitle.TabIndex = 2;
            lblTitle.Text = "Lista Facturi";
            // 
            // panelBottom
            // 
            panelBottom.Controls.Add(btnGenereazaWord);
            panelBottom.Dock = DockStyle.Bottom;
            panelBottom.Location = new Point(0, 390);
            panelBottom.Name = "panelBottom";
            panelBottom.Size = new Size(800, 60);
            panelBottom.TabIndex = 3;
            // 
            // ListaFacturiForm
            // 
            AutoScaleDimensions = new SizeF(7F, 16F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(panelBottom);
            Controls.Add(lblTitle);
            Controls.Add(dgvFacturi);
            Font = new Font("Yu Gothic", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Name = "ListaFacturiForm";
            Text = "Lista Facturi";
            ((System.ComponentModel.ISupportInitialize)dgvFacturi).EndInit();
            panelBottom.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dgvFacturi;
        private Button btnGenereazaWord;
        private Label lblTitle;
        private Panel panelBottom;
    }
}
