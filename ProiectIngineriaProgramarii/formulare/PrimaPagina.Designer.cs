namespace ProiectIngineriaProgramarii
{
    partial class StartForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            menuStrip1 = new MenuStrip();
            clientiToolStripMenuItem = new ToolStripMenuItem();
            produseToolStripMenuItem = new ToolStripMenuItem();
            facturiToolStripMenuItem = new ToolStripMenuItem();
            adaugaFacturaToolStripMenuItem = new ToolStripMenuItem();
            listaFacturiToolStripMenuItem = new ToolStripMenuItem();
            rapoarteToolStripMenuItem = new ToolStripMenuItem();
            lblWelcome = new Label();
            lblDescription = new Label();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { clientiToolStripMenuItem, produseToolStripMenuItem, facturiToolStripMenuItem, rapoarteToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // clientiToolStripMenuItem
            // 
            clientiToolStripMenuItem.Name = "clientiToolStripMenuItem";
            clientiToolStripMenuItem.Size = new Size(56, 20);
            clientiToolStripMenuItem.Text = "Clienti";
            // 
            // produseToolStripMenuItem
            // 
            produseToolStripMenuItem.Name = "produseToolStripMenuItem";
            produseToolStripMenuItem.Size = new Size(64, 20);
            produseToolStripMenuItem.Text = "Produse";
            // 
            // facturiToolStripMenuItem
            // 
            facturiToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { adaugaFacturaToolStripMenuItem, listaFacturiToolStripMenuItem });
            facturiToolStripMenuItem.Name = "facturiToolStripMenuItem";
            facturiToolStripMenuItem.Size = new Size(56, 20);
            facturiToolStripMenuItem.Text = "Facturi";
            // 
            // adaugaFacturaToolStripMenuItem
            // 
            adaugaFacturaToolStripMenuItem.Name = "adaugaFacturaToolStripMenuItem";
            adaugaFacturaToolStripMenuItem.Size = new Size(180, 22);
            adaugaFacturaToolStripMenuItem.Text = "Adauga factura";
            // 
            // listaFacturiToolStripMenuItem
            // 
            listaFacturiToolStripMenuItem.Name = "listaFacturiToolStripMenuItem";
            listaFacturiToolStripMenuItem.Size = new Size(180, 22);
            listaFacturiToolStripMenuItem.Text = "Lista facturi";
            // 
            // rapoarteToolStripMenuItem
            // 
            rapoarteToolStripMenuItem.Name = "rapoarteToolStripMenuItem";
            rapoarteToolStripMenuItem.Size = new Size(68, 20);
            rapoarteToolStripMenuItem.Text = "Rapoarte";
            // 
            // lblWelcome
            // 
            lblWelcome.AutoSize = true;
            lblWelcome.Font = new Font("Yu Gothic", 24F, FontStyle.Bold);
            lblWelcome.Location = new Point(200, 150);
            lblWelcome.Name = "lblWelcome";
            lblWelcome.Size = new Size(400, 42);
            lblWelcome.TabIndex = 1;
            lblWelcome.Text = "Sistem de Gestiune";
            lblWelcome.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // lblDescription
            // 
            lblDescription.AutoSize = true;
            lblDescription.Font = new Font("Yu Gothic", 12F);
            lblDescription.Location = new Point(180, 220);
            lblDescription.Name = "lblDescription";
            lblDescription.Size = new Size(440, 21);
            lblDescription.TabIndex = 2;
            lblDescription.Text = "Selecteaza o optiune din meniul de sus pentru a incepe";
            lblDescription.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // StartForm
            // 
            AutoScaleDimensions = new SizeF(7F, 16F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(lblDescription);
            Controls.Add(lblWelcome);
            Controls.Add(menuStrip1);
            Font = new Font("Yu Gothic", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            MainMenuStrip = menuStrip1;
            Name = "StartForm";
            Text = "Sistem Gestiune";
            Load += Form1_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem clientiToolStripMenuItem;
        private ToolStripMenuItem produseToolStripMenuItem;
        private ToolStripMenuItem facturiToolStripMenuItem;
        private ToolStripMenuItem adaugaFacturaToolStripMenuItem;
        private ToolStripMenuItem listaFacturiToolStripMenuItem;
        private ToolStripMenuItem rapoarteToolStripMenuItem;
        private Label lblWelcome;
        private Label lblDescription;
    }
}
