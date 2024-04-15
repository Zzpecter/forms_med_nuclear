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
    public partial class FrmHema : Form
    {
        private DataTable dtValores, dtImpresion;
        BLL.Hema h;

        public FrmHema()
        {
            InitializeComponent();
            dtValores = new DataTable();
            dtImpresion = new DataTable();
            h = new BLL.Hema();
        }

        private void FrmHema_Load(object sender, EventArgs e)
        {
            Cargar();
        }

        private void Cargar()
        {
            dtValores = h.Leer(@"E:\Examenes\Referencias\hematologia.xlsx");

            //Carga unidades
            lblU01.Text = dtValores.Rows[0].ItemArray[0].ToString();
            lblU02.Text = dtValores.Rows[1].ItemArray[0].ToString();
            lblU03.Text = dtValores.Rows[2].ItemArray[0].ToString();
            lblU04.Text = dtValores.Rows[3].ItemArray[0].ToString();
            lblU05.Text = dtValores.Rows[4].ItemArray[0].ToString();
            lblU06.Text = dtValores.Rows[5].ItemArray[0].ToString();
            lblU07.Text = dtValores.Rows[6].ItemArray[0].ToString();
            lblU08.Text = dtValores.Rows[7].ItemArray[0].ToString();
            lblU09.Text = dtValores.Rows[8].ItemArray[0].ToString();
            lblU10.Text = dtValores.Rows[9].ItemArray[0].ToString();
            lblU11.Text = dtValores.Rows[10].ItemArray[0].ToString();
            lblU12.Text = dtValores.Rows[11].ItemArray[0].ToString();
            lblU13.Text = dtValores.Rows[12].ItemArray[0].ToString();
            lblU14.Text = dtValores.Rows[13].ItemArray[0].ToString();
            lblU15.Text = dtValores.Rows[14].ItemArray[0].ToString();

            //Carga referencias
            lblR01.Text = dtValores.Rows[0].ItemArray[1].ToString();
            lblR02.Text = dtValores.Rows[1].ItemArray[1].ToString();
            lblR03.Text = dtValores.Rows[2].ItemArray[1].ToString();
            lblR04.Text = dtValores.Rows[3].ItemArray[1].ToString();
            lblR05.Text = dtValores.Rows[4].ItemArray[1].ToString();
            lblR06.Text = dtValores.Rows[5].ItemArray[1].ToString();
            lblR07.Text = dtValores.Rows[6].ItemArray[1].ToString();
            lblR08.Text = dtValores.Rows[7].ItemArray[1].ToString();
            lblR09.Text = dtValores.Rows[8].ItemArray[1].ToString();
            lblR10.Text = dtValores.Rows[9].ItemArray[1].ToString();
            lblR11.Text = dtValores.Rows[10].ItemArray[1].ToString();
            lblR12.Text = dtValores.Rows[11].ItemArray[1].ToString();
            lblR13.Text = dtValores.Rows[12].ItemArray[1].ToString();
            lblR14.Text = dtValores.Rows[13].ItemArray[1].ToString();
            lblR15.Text = dtValores.Rows[14].ItemArray[1].ToString();

            
            dtImpresion.Columns.Add("Nombre");
            dtImpresion.Columns.Add("Valor");
            dtImpresion.Columns.Add("Unidad");
            dtImpresion.Columns.Add("Referencia");
        }

        private void btnVolver_Click(object sender, EventArgs e)
        {
            FrmMainMenu f = new FrmMainMenu();
            f.Show();
            this.Hide();
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            if (tbDoctor.Text.Equals(string.Empty) || tbNombre.Text.Equals(string.Empty))
                MessageBox.Show("Ingrese nombre del Paciente y del doctor primero!");
            else
            {
                CrearTabla();
                string dir = h.Escribir(dtImpresion, tbNombre.Text, tbDoctor.Text, DateTime.Now);
                h.PrintExcel(dir);
                FrmMainMenu f = new FrmMainMenu();
                f.Show();
                this.Hide();
            }
        }

        private void CrearTabla()
        {
            if(!tbV01.Text.Equals(string.Empty) && !tbV01.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Eritrocitos", tbV01.Text, lblU01.Text, lblR01.Text);

            if(!tbV02.Text.Equals(string.Empty) && !tbV02.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Hemoglobina", tbV02.Text, lblU02.Text, lblR02.Text);

            if(!tbV03.Text.Equals(string.Empty) && !tbV03.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Hematocrito", tbV03.Text, lblU03.Text, lblR03.Text);

            if(!tbV04.Text.Equals(string.Empty) && !tbV04.Text.Equals(" ")) 
                dtImpresion.Rows.Add("VSG 1a Hora", tbV04.Text, lblU04.Text, lblR04.Text);

            if(!tbV05.Text.Equals(string.Empty) && !tbV05.Text.Equals(" ")) 
                dtImpresion.Rows.Add("VSG 2a Hora", tbV05.Text, lblU05.Text, lblR05.Text);

            if(!tbV06.Text.Equals(string.Empty) && !tbV06.Text.Equals(" ")) 
                dtImpresion.Rows.Add("T. Coagulación", tbV06.Text, lblU06.Text, lblR06.Text);

            if(!tbV07.Text.Equals(string.Empty) && !tbV07.Text.Equals(" ")) 
                dtImpresion.Rows.Add("T. Sangría", tbV07.Text, lblU07.Text, lblR07.Text);

            if(!tbV08.Text.Equals(string.Empty) && !tbV08.Text.Equals(" ")) 
                dtImpresion.Rows.Add("T. Protrombina", tbV08.Text, lblU08.Text, lblR08.Text);

            if(!tbV09.Text.Equals(string.Empty) && !tbV09.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Leucocitos", tbV09.Text, lblU09.Text, lblR09.Text);

            if(!tbV10.Text.Equals(string.Empty) && !tbV10.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Basófilos", tbV10.Text, lblU10.Text, lblR10.Text);

            if(!tbV11.Text.Equals(string.Empty) && !tbV11.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Eosinófilos", tbV11.Text, lblU11.Text, lblR11.Text);

            if(!tbV12.Text.Equals(string.Empty) && !tbV12.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Bastonados", tbV12.Text, lblU12.Text, lblR12.Text);

            if(!tbV13.Text.Equals(string.Empty) && !tbV13.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Segmentados", tbV13.Text, lblU13.Text, lblR13.Text);

            if(!tbV14.Text.Equals(string.Empty) && !tbV14.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Linfocitos", tbV14.Text, lblU14.Text, lblR14.Text);

            if(!tbV15.Text.Equals(string.Empty) && !tbV15.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Monocitos", tbV15.Text, lblU15.Text, lblR15.Text);

            if(!tbV16.Text.Equals(string.Empty) && !tbV16.Text.Equals(" ")) 
                dtImpresion.Rows.Add("Grupo Sanguíneo", tbV16.Text, "-", "-");

        }

        private void FrmHema_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
