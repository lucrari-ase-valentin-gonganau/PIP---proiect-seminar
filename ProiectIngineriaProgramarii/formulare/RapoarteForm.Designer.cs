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
            this.btnVeziGraficExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnVeziGraficExcel
            // 
            this.btnVeziGraficExcel.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.btnVeziGraficExcel.Location = new System.Drawing.Point(250, 150);
            this.btnVeziGraficExcel.Name = "btnVeziGraficExcel";
            this.btnVeziGraficExcel.Size = new System.Drawing.Size(300, 80);
            this.btnVeziGraficExcel.TabIndex = 0;
            this.btnVeziGraficExcel.Text = "Vezi Graficul în Excel";
            this.btnVeziGraficExcel.UseVisualStyleBackColor = true;
            this.btnVeziGraficExcel.Click += new System.EventHandler(this.btnVeziGraficExcel_Click);
            // 
            // RapoarteForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnVeziGraficExcel);
            this.Name = "RapoarteForm";
            this.Text = "RapoarteForm";
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Button btnVeziGraficExcel;
    }
}