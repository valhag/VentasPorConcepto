using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MyExcel = Microsoft.Office.Interop.Excel; 

namespace VentasPorConcepto
{
    public partial class Form2 : Form
    {
        Class1 x = new Class1();
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.Text = " REPORTE FACTURAS/PAGOS";
            string mensaje;
            this.comboBox1.Items.Clear();
            this.comboBox1.DataSource = x.mCargarEmpresas(out mensaje);
            comboBox1.DisplayMember = "Nombre";
            comboBox1.ValueMember = "Ruta";
            try
            {
                this.comboBox1.SelectedIndex = 1;
                this.comboBox1.SelectedIndex = 0;
            }
            catch (Exception ee)
            {
                this.comboBox1.SelectedIndex = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            DataTable Datos = null;
            string lquery;
            lquery = " select m6.cnombrec01, m8.cfolio, m8.cseriedo01, m8.ctotal, m8.creferen01, m8.cpendiente, m8a.cseriedo01, m8a.cfolio, m9.cimporte01, m6a.cnombrec01 " +
" from mgw10008 m8  " +
" join mgw10009 m9 on m9.ciddocum02 = m8.ciddocum01  " +
" join mgw10006 m6 on m6.cidconce01 = m8.cidconce01  " +
" join mgw10008 m8a on m9.ciddocum01 = m8a.ciddocum01  " +
" join mgw10006 m6a on m8a.cidconce01 = m6a.cidconce01  " +
" where m8.ciddocum02 = 4 and m8.ccancelado = 0  " +
            " and dtos(m8.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "'" ;

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                lquery += " and m8.cfolio between " + textBox1.Text + " and " + textBox2.Text;
            }

            string sconceptos = " and m8.cidconce01 in (";
            List<RegConcepto> misconceptos = new List<RegConcepto>();
            for (int i = 0; i < listBox1.SelectedItems.Count; i++)
            {
                RegConcepto v = new RegConcepto();
                long zz = ((VentasPorConcepto.Class1.RegConcepto)(listBox1.SelectedItems[i])).id;
                sconceptos += zz.ToString();
                sconceptos += ",";
            }  

            if (sconceptos == " and m8.cidconce01 in (")
                sconceptos = "";
            else
            {
                sconceptos = sconceptos.Substring(0, sconceptos.Length - 1);
                sconceptos += ")";
            }
            lquery += sconceptos;
             lquery += " order by m8.cfecha, m8.cfolio";
            x.mTraerInformacionPrimerReporte(lquery, comboBox1.SelectedValue.ToString());

            x.mReporteFacturaAbono(comboBox1.Text    );


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                List<RegConcepto> _listaConceptosFacturaOrigen = new List<RegConcepto>();
                //listBox1.Items.Clear();
                listBox1.DataSource = x.mCargarConceptos(comboBox1.SelectedValue.ToString());
                listBox1.DisplayMember = "Nombre";
                listBox1.ValueMember = "id";

            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }
    }
}
