using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VentasPorConcepto
{
    public partial class OC : Form
    {
        Class1 x = new Class1();
        public OC()
        {
            InitializeComponent();
            catalogo1.setLabel("Proveedor Ini.");
            catalogo2.setLabel("Proveedor Fin");
        }

        private void OC_Load(object sender, EventArgs e)
        {
            this.Text = " REPORTE Ordenes de Compra";
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

            string lquery;



            lquery = "select dtos(m8.cfecha) as cfecha, m8.cfolio, m2.crazonso01, m5.cnombrep01, m10.cunidades, m10.cunidade03, m5.ccodigop01, m5.ccodaltern, m5.cnomaltern from mgw10008 m8 " +
" join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
" join mgw10010 m10 on m8.ciddocum01 = m10.ciddocum01 " +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01 " +
" where m8.ciddocum02 =17 " +
" and m8.ccancelado = 0 " +
" and dtos(m8.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' and m8.ccancelado = 0   ";



    if (catalogo1.mRegresarCodigo() != "" && catalogo2.mRegresarCodigo() != "")
        lquery += " and m2.ccodigoc01 >= '" + catalogo1.mRegresarCodigo() + "' and m2.ccodigoc01 <= '" + catalogo2.mRegresarCodigo() + "'";



lquery += " order by m8.cfecha, m8.cfolio";


            x.mTraerInformacionPrimerReporte(lquery, comboBox1.SelectedValue.ToString());

            string lfechai = dateTimePicker1.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Year.ToString().PadLeft(4, '0');
            string lfechaf = dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Year.ToString().PadLeft(4, '0');


            x.mReporteOC(comboBox1.Text, lfechai, lfechaf);
        }
    }
}
