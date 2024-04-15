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
    public partial class FrmRadio : Form
    {
        private DataTable dtValores, dtImpresion;
        private BLL.Radio r;

        public FrmRadio()
        {
            InitializeComponent();
            dtValores = new DataTable();
            dtImpresion = new DataTable();
            r = new BLL.Radio();
        }

        private void FrmRadio_Load(object sender, EventArgs e)
        {
            Cargar();
        }

        private void Cargar()
        {
            dtValores = r.Leer(@"E:\Examenes\Referencias\radio.xlsx");

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

            //Carga referencias
            lblR01.Text = dtValores.Rows[0].ItemArray[1].ToString();
            lblR02.Text = dtValores.Rows[1].ItemArray[1].ToString();
            lblR03.Text = dtValores.Rows[2].ItemArray[1].ToString();
            lblR04.Text = dtValores.Rows[3].ItemArray[1].ToString();
            lblR06.Text = dtValores.Rows[5].ItemArray[1].ToString();
            lblR07.Text = dtValores.Rows[6].ItemArray[1].ToString();
            lblR08.Text = dtValores.Rows[7].ItemArray[1].ToString();
            lblR09.Text = dtValores.Rows[8].ItemArray[1].ToString();
            lblR10.Text = dtValores.Rows[9].ItemArray[1].ToString();
            lblR11.Text = dtValores.Rows[10].ItemArray[1].ToString();
            lblR12.Text = dtValores.Rows[11].ItemArray[1].ToString();
            lblR18.Text = dtValores.Rows[17].ItemArray[1].ToString();
            lblR19.Text = dtValores.Rows[18].ItemArray[1].ToString();
            lblR20.Text = dtValores.Rows[19].ItemArray[1].ToString();
            lblR21.Text = dtValores.Rows[20].ItemArray[1].ToString();


            dtImpresion.Columns.Add("Nombre");
            dtImpresion.Columns.Add("Valor");
            dtImpresion.Columns.Add("Unidad");
            dtImpresion.Columns.Add("Referencia");
        }

        private void lblR05_MouseMove(object sender, MouseEventArgs e)
        {
            tT01.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[4].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT01.SetToolTip(lblR05, s);
        }

        private void lblR13_MouseMove(object sender, MouseEventArgs e)
        {
            tT02.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[12].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT02.SetToolTip(lblR13, s);
        }

        private void lblR14_MouseMove(object sender, MouseEventArgs e)
        {
            tT03.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[13].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT03.SetToolTip(lblR14, s);
        }

        private void lblR15_MouseMove(object sender, MouseEventArgs e)
        {
            tT04.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[14].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT04.SetToolTip(lblR15, s);
        }

        private void lblR16_MouseMove(object sender, MouseEventArgs e)
        {
            tT05.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[15].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT05.SetToolTip(lblR16, s);
        }

        private void lblR17_MouseMove(object sender, MouseEventArgs e)
        {
            tT06.ToolTipTitle = "Referencia";
            string s = dtValores.Rows[16].ItemArray[1].ToString();
            s = s.Replace("@", System.Environment.NewLine);
            tT06.SetToolTip(lblR17, s);
        }

        private void FrmRadio_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            if (tbDoctor.Text.Equals(string.Empty) || tbNombre.Text.Equals(string.Empty))
                MessageBox.Show("Ingrese nombre del Paciente y del doctor primero!");
            else
            {
                CrearTabla();
                string dir = r.Escribir(dtImpresion, tbNombre.Text, tbDoctor.Text, DateTime.Now);
                r.PrintExcel(dir);
                FrmMainMenu f = new FrmMainMenu();
                f.Show();
                this.Hide();
            }
        }

        private void btnVolver_Click(object sender, EventArgs e)
        {
            FrmMainMenu f = new FrmMainMenu();
            f.Show();
            this.Hide();
        }

        private void CrearTabla()
        {
            if (!tbV01.Text.Equals(string.Empty) && !tbV01.Text.Equals(" "))
                dtImpresion.Rows.Add("T4", tbV01.Text, lblU01.Text, lblR01.Text);

            if (!tbV02.Text.Equals(string.Empty) && !tbV02.Text.Equals(" "))
                dtImpresion.Rows.Add("T4", tbV02.Text, lblU02.Text, lblR02.Text);

            if (!tbV03.Text.Equals(string.Empty) && !tbV03.Text.Equals(" "))
                dtImpresion.Rows.Add("T3", tbV03.Text, lblU03.Text, lblR03.Text);

            if (!tbV04.Text.Equals(string.Empty) && !tbV04.Text.Equals(" "))
                dtImpresion.Rows.Add("T3", tbV04.Text, lblU04.Text, lblR04.Text);

            if (!tbV05.Text.Equals(string.Empty) && !tbV05.Text.Equals(" "))
            {
                string[] refs = dtValores.Rows[4].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add("TSH", tbV05.Text, lblU05.Text, refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[2]);
            }

            if (!tbV06.Text.Equals(string.Empty) && !tbV06.Text.Equals(" "))
                dtImpresion.Rows.Add("T4 Libre", tbV06.Text, lblU06.Text, lblR06.Text);

            if (!tbV07.Text.Equals(string.Empty) && !tbV07.Text.Equals(" "))
                dtImpresion.Rows.Add("PRL", tbV07.Text, lblU07.Text, lblR07.Text);

            if (!tbV08.Text.Equals(string.Empty) && !tbV08.Text.Equals(" "))
                dtImpresion.Rows.Add("HGH:(H. Crecimiento Basal)", tbV08.Text, lblU08.Text, lblR08.Text);

            if (!tbV09.Text.Equals(string.Empty) && !tbV09.Text.Equals(" "))
                dtImpresion.Rows.Add("HGH:(H. Crecimiento Post Ejercicio)", tbV09.Text, lblU09.Text, lblR09.Text);

            if (!tbV10.Text.Equals(string.Empty) && !tbV10.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("HGH:(H. Crecimiento Post Estímulo con", tbV10.Text, lblU10.Text, " ");
                dtImpresion.Rows.Add("de", tbV11.Text, " ", " ");
                dtImpresion.Rows.Add("mas", tbV12.Text, lblU11.Text, " ");
                dtImpresion.Rows.Add("minutos de actividad física)", tbV13.Text, lblU12.Text, lblR10.Text);
            }

            if (!tbV14.Text.Equals(string.Empty) && !tbV14.Text.Equals(" "))
                dtImpresion.Rows.Add("CORTISOL: a.m.", tbV14.Text, lblU13.Text, lblR11.Text);

            if (!tbV15.Text.Equals(string.Empty) && !tbV15.Text.Equals(" "))
                dtImpresion.Rows.Add("CORTISOL: p.m.", tbV15.Text, lblU14.Text, lblR12.Text);

            if (!tbV16.Text.Equals(string.Empty) && !tbV16.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("LH", tbV16.Text, lblU15.Text, " ");
                dtImpresion.Rows.Add("Día de Ciclo", tbV17.Text, " ", " ");
                string[] refs = dtValores.Rows[12].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add(" ", " ", " ", refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[2]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[3]);
            }

            if (!tbV18.Text.Equals(string.Empty) && !tbV18.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("FSH", tbV18.Text, lblU16.Text, " ");
                dtImpresion.Rows.Add("Día de Ciclo", tbV19.Text, " ", " ");
                string[] refs = dtValores.Rows[13].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add(" ", " ", " ", refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[2]);
            }

            if (!tbV20.Text.Equals(string.Empty) && !tbV20.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("ESTRADIOL", tbV20.Text, lblU17.Text, " ");
                dtImpresion.Rows.Add("Día de Ciclo", tbV21.Text, " ", " ");
                string[] refs = dtValores.Rows[14].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add(" ", " ", " ", refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[2]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[3]);
            }

            if (!tbV22.Text.Equals(string.Empty) && !tbV22.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("TESTOSTERONA", tbV22.Text, lblU18.Text, " ");
                string[] refs = dtValores.Rows[15].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add(" ", " ", " ", refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[2]);
            }

            if (!tbV23.Text.Equals(string.Empty) && !tbV23.Text.Equals(" "))
            {
                dtImpresion.Rows.Add("PROGESTERONA", tbV23.Text, lblU19.Text, " ");
                string[] refs = dtValores.Rows[16].ItemArray[1].ToString().Split('@');
                dtImpresion.Rows.Add(" ", " ", " ", refs[0]);
                dtImpresion.Rows.Add(" ", " ", " ", refs[1]);
            }

            if (!tbV24.Text.Equals(string.Empty) && !tbV24.Text.Equals(" "))
                dtImpresion.Rows.Add("INSULINA: Basal", tbV24.Text, lblU20.Text, lblR18.Text);

            if (!tbV25.Text.Equals(string.Empty) && !tbV25.Text.Equals(" "))
                dtImpresion.Rows.Add("INSULINA: Postprandial", tbV25.Text, lblU21.Text, lblR19.Text);

            if (!tbV26.Text.Equals(string.Empty) && !tbV26.Text.Equals(" "))
                dtImpresion.Rows.Add("Antígeno Prostático Total", tbV26.Text, lblU22.Text, lblR20.Text);

            if (!tbV27.Text.Equals(string.Empty) && !tbV27.Text.Equals(" "))
                dtImpresion.Rows.Add("Antígeno Prostático Libre", tbV27.Text, lblU23.Text, lblR21.Text);
        }
    }
}
