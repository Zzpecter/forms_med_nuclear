using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GUI
{
    public partial class FrmMainMenu : Form
    {
        public FrmMainMenu()
        {
            InitializeComponent();
        }

        private void btnHema_Click(object sender, EventArgs e)
        {
            FrmHema f = new FrmHema();
            f.Show();
            this.Hide();
        }

        private void btnQuimica_Click(object sender, EventArgs e)
        {
            FrmQuimica f = new FrmQuimica();
            f.Show();
            this.Hide();
        }

        private void btnRadio_Click(object sender, EventArgs e)
        {
            FrmRadio f = new FrmRadio();
            f.Show();
            this.Hide();
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FrmMainMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
