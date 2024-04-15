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
    public partial class FrmQuimica : Form
    {
        private DataTable dtValores, dtImpresion;
        private BLL.Quimica q;

        public FrmQuimica()
        {
            InitializeComponent();
            dtValores = new DataTable();
            dtImpresion = new DataTable();
            q = new BLL.Quimica();
        }

        private void FrmQuimica_Load(object sender, EventArgs e)
        {
            Cargar();
        }

        private void Cargar()
        {
            dtValores = q.Leer(@"E:\Examenes\Referencias\quimica.xlsx");
            string s = string.Empty;

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
            lblU16.Text = dtValores.Rows[15].ItemArray[0].ToString();
            lblU17.Text = dtValores.Rows[16].ItemArray[0].ToString();
            lblU18.Text = dtValores.Rows[17].ItemArray[0].ToString();
            lblU19.Text = dtValores.Rows[18].ItemArray[0].ToString();
            lblU20.Text = dtValores.Rows[19].ItemArray[0].ToString();
            lblU21.Text = dtValores.Rows[20].ItemArray[0].ToString();
            lblU22.Text = dtValores.Rows[21].ItemArray[0].ToString();
            lblU23.Text = dtValores.Rows[22].ItemArray[0].ToString();
            lblU24.Text = dtValores.Rows[23].ItemArray[0].ToString();

            //Carga referencias
            lblR01.Text = dtValores.Rows[0].ItemArray[1].ToString();
            lblR02.Text = dtValores.Rows[1].ItemArray[1].ToString();
            lblR03.Text = dtValores.Rows[2].ItemArray[1].ToString();
            lblR04.Text = dtValores.Rows[3].ItemArray[1].ToString();
            lblR05.Text = dtValores.Rows[4].ItemArray[1].ToString();
            s = dtValores.Rows[5].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            lblR06.Text = s;
            s = dtValores.Rows[6].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            lblR07.Text = s;
            lblR08.Text = dtValores.Rows[7].ItemArray[1].ToString();
            lblR09.Text = dtValores.Rows[8].ItemArray[1].ToString();
            lblR10.Text = dtValores.Rows[9].ItemArray[1].ToString();
            lblR11.Text = dtValores.Rows[10].ItemArray[1].ToString();
            lblR12.Text = dtValores.Rows[11].ItemArray[1].ToString();
            lblR13.Text = dtValores.Rows[12].ItemArray[1].ToString();
            lblR14.Text = dtValores.Rows[13].ItemArray[1].ToString();
            s = dtValores.Rows[14].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            lblR15.Text = s;
            lblR16.Text = dtValores.Rows[15].ItemArray[1].ToString();
            lblR17.Text = dtValores.Rows[16].ItemArray[1].ToString();
            lblR18.Text = dtValores.Rows[17].ItemArray[1].ToString();
            lblR19.Text = dtValores.Rows[18].ItemArray[1].ToString();
            lblR20.Text = dtValores.Rows[19].ItemArray[1].ToString();
            lblR21.Text = dtValores.Rows[20].ItemArray[1].ToString();
            lblR22.Text = dtValores.Rows[21].ItemArray[1].ToString();
            lblR23.Text = dtValores.Rows[22].ItemArray[1].ToString();
            lblR24.Text = dtValores.Rows[23].ItemArray[1].ToString();


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
                string dir = q.Escribir(dtImpresion, tbNombre.Text, tbDoctor.Text, DateTime.Now);
                q.PrintExcel(dir);
                FrmMainMenu f = new FrmMainMenu();
                f.Show();
                this.Hide();
            }
        }

        private void CrearTabla()
        {
            if (!tbV01.Text.Equals(string.Empty) && !tbV01.Text.Equals(" "))
                dtImpresion.Rows.Add("Glucosa Basal", tbV01.Text, lblU01.Text, lblR01.Text);

            if (!tbV02.Text.Equals(string.Empty) && !tbV02.Text.Equals(" "))
                dtImpresion.Rows.Add("Glucosa Postprandial", tbV02.Text, lblU02.Text, lblR02.Text);

            if (!tbV03.Text.Equals(string.Empty) && !tbV03.Text.Equals(" "))
                dtImpresion.Rows.Add("Hb Glicosilada (Hb A1)", tbV03.Text, lblU03.Text, lblR03.Text);

            if (!tbV04.Text.Equals(string.Empty) && !tbV04.Text.Equals(" "))
                dtImpresion.Rows.Add("Colesterol", tbV04.Text, lblU04.Text, lblR04.Text);

            if (!tbV05.Text.Equals(string.Empty) && !tbV05.Text.Equals(" "))
                dtImpresion.Rows.Add("HDL Col", tbV05.Text, lblU05.Text, lblR05.Text);

            if (!tbV06.Text.Equals(string.Empty) && !tbV06.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("LDL Col", tbV06.Text, lblU06.Text, "riesgo bajo o nulo menor a 140");
                dtImpresion.Rows.Add(" ", " ", " ", "riesgo moderado de 140 a 190");
                dtImpresion.Rows.Add(" ", " ", " ", "riesgo elevado mayor a 190");
            }

            if (!tbV07.Text.Equals(string.Empty) && !tbV07.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("Triglicéridos", tbV07.Text, lblU07.Text, "0,10 - 1,50");
                dtImpresion.Rows.Add(" ", " ", " ", "Diabéticos con 1 factor riesgo hasta 1,30");
                dtImpresion.Rows.Add(" ", " ", " ", "Diabéticos con mas de 2 factores riesgo hasta 1,10");
            }

            if (!tbV08.Text.Equals(string.Empty) && !tbV08.Text.Equals(" "))
                dtImpresion.Rows.Add("Bilirrubina Total", tbV08.Text, lblU08.Text, lblR08.Text);

            if (!tbV09.Text.Equals(string.Empty) && !tbV09.Text.Equals(" "))
                dtImpresion.Rows.Add("Bilirrubina Directa", tbV09.Text, lblU09.Text, lblR09.Text);

            if (!tbV10.Text.Equals(string.Empty) && !tbV10.Text.Equals(" "))
                dtImpresion.Rows.Add("Bilirrubina Indirecta", tbV10.Text, lblU10.Text, lblR10.Text);

            if (!tbV11.Text.Equals(string.Empty) && !tbV11.Text.Equals(" "))
                dtImpresion.Rows.Add("GOT", tbV11.Text, lblU11.Text, lblR11.Text);

            if (!tbV12.Text.Equals(string.Empty) && !tbV12.Text.Equals(" "))
                dtImpresion.Rows.Add("GPT", tbV12.Text, lblU12.Text, lblR12.Text);

            if (!tbV13.Text.Equals(string.Empty) && !tbV13.Text.Equals(" "))
                dtImpresion.Rows.Add("Fosfatasa Alcalina", tbV13.Text, lblU13.Text, lblR13.Text);

            if (!tbV14.Text.Equals(string.Empty) && !tbV14.Text.Equals(" "))
                dtImpresion.Rows.Add("Urea", tbV14.Text, lblU14.Text, lblR14.Text);

            if (!tbV15.Text.Equals(string.Empty) && !tbV15.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("Ácido Úrico", tbV15.Text, lblU15.Text, "varones 25 - 60");
                dtImpresion.Rows.Add(" ", " ", " ", "mujeres 20 - 50");

            }

            if (!tbV16.Text.Equals(string.Empty) && !tbV16.Text.Equals(" "))
                dtImpresion.Rows.Add("Creatinina", tbV16.Text, lblU16.Text, lblR16.Text);

            if (!tbV17.Text.Equals(string.Empty) && !tbV17.Text.Equals(" "))
                dtImpresion.Rows.Add("Calcio Sérico", tbV17.Text, lblU17.Text, lblR17.Text);

            if (!tbV18.Text.Equals(string.Empty) || !tbV19.Text.Equals(string.Empty) || !tbV20.Text.Equals(string.Empty) || !tbV21.Text.Equals(string.Empty) || !tbV22.Text.Equals(string.Empty) || !tbV23.Text.Equals(string.Empty) || !tbV24.Text.Equals(string.Empty))
                dtImpresion.Rows.Add("Tolerancia a la Glucosa", " ", " ", " ");

            if (!tbV18.Text.Equals(string.Empty) && !tbV18.Text.Equals(" "))
                dtImpresion.Rows.Add("0 minutos", tbV18.Text, lblU18.Text, lblR18.Text);

            if (!tbV19.Text.Equals(string.Empty) && !tbV19.Text.Equals(" "))
                dtImpresion.Rows.Add("30 minutos", tbV19.Text, lblU19.Text, lblR19.Text);

            if (!tbV20.Text.Equals(string.Empty) && !tbV20.Text.Equals(" "))
                dtImpresion.Rows.Add("60 minutos", tbV20.Text, lblU20.Text, lblR20.Text);

            if (!tbV21.Text.Equals(string.Empty) && !tbV21.Text.Equals(" "))
                dtImpresion.Rows.Add("90 minutos", tbV21.Text, lblU21.Text, lblR21.Text);

            if (!tbV22.Text.Equals(string.Empty) && !tbV22.Text.Equals(" "))
                dtImpresion.Rows.Add("120 minutos", tbV22.Text, lblU22.Text, lblR22.Text);

            if (!tbV23.Text.Equals(string.Empty) && !tbV23.Text.Equals(" "))
                dtImpresion.Rows.Add("150 minutos", tbV23.Text, lblU23.Text, lblR23.Text);

            if (!tbV24.Text.Equals(string.Empty) && !tbV24.Text.Equals(" "))
                dtImpresion.Rows.Add("180 minutos", tbV24.Text, lblU24.Text, lblR24.Text);
        }

        private void FrmQuimica_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
