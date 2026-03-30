namespace ProiectIngineriaProgramarii
{
    partial class RapoarteForm
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
            btnRaportInteractiv = new Button();
            SuspendLayout();
            // 
            // btnRaportInteractiv
            // 
            btnRaportInteractiv.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnRaportInteractiv.Location = new Point(219, 112);
            btnRaportInteractiv.Margin = new Padding(3, 2, 3, 2);
            btnRaportInteractiv.Name = "btnRaportInteractiv";
            btnRaportInteractiv.Size = new Size(262, 60);
            btnRaportInteractiv.TabIndex = 0;
            btnRaportInteractiv.Text = "Vezi Graficul In Excel (Cu Evenimente)";
            btnRaportInteractiv.UseVisualStyleBackColor = true;
            btnRaportInteractiv.Click += btnRaportInteractiv_Click;
            // 
            // RapoarteForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(700, 338);
            Controls.Add(btnRaportInteractiv);
            Margin = new Padding(3, 2, 3, 2);
            Name = "RapoarteForm";
            Text = "RapoarteForm";
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Button btnRaportInteractiv;
    }
}