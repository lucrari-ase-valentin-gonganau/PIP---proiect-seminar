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

        public RapoarteForm(StartForm mainForm)
        {
            InitializeComponent();
            this.Text = "Rapoarte";
            _mainForm = mainForm;
        }

        public RapoarteForm()
        {
            InitializeComponent();
            this.Text = "Rapoarte";
        }
    }
}
