using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ProiectIngineriaProgramarii
{
    public partial class RapoarteForm : Form
    {
        private StartForm _mainForm;
        private Button btnBack;

        public RapoarteForm(StartForm mainForm)
        {
            InitializeComponent();
            _mainForm = mainForm;
            AddBackButton();
        }

        public RapoarteForm()
        {
            InitializeComponent();
        }

        private void AddBackButton()
        {
            btnBack = new Button
            {
                Text = "← Înapoi",
                Location = new Point(10, 10),
                Size = new Size(100, 35),
                Font = new Font("Yu Gothic", 10F, FontStyle.Bold),
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                TabIndex = 999
            };
            btnBack.FlatAppearance.BorderSize = 0;
            btnBack.Click += (s, e) => _mainForm?.ShowMainMenu();
            this.Controls.Add(btnBack);
            btnBack.BringToFront();
        }
    }
}
