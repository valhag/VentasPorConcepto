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
    public partial class IEPScs : Form
    {
        Class1 x = new Class1();
        public IEPScs()
        {
            InitializeComponent();
        }

        private void IEPScs_Load(object sender, EventArgs e)
        {
            this.Text = " REPORTE IEPS";
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



            lquery = "SELECT m8.cfolio as foliocargo, m8.cseriedo01 as seriecargo, m8.cfecha as fechacargo, m2.crazonso01 as cliente, m6.cnombrec01 as conceptocargo, " +
" m8.cneto as netocargo, m8.cimpuesto1 as iva, m8.cimpuesto2 as ieps, m8.ctotal as totalcargo, " +
" m8a.cfolio as folioabono, m8a.cseriedo01 as serieabono, m8a.cfecha as fechaabono, m6a.cnombrec01 as conceptoabono, m8a.ctotal as totalabono, m9.cimporte01 as pagado, m8a.creferen01 as referencia, m8a.cobserva01 as observaciones" +
" FROM MGW10008 m8 join mgw10002 m2 " +
" on m8.cidclien01 = m2.cidclien01 " +
" join mgw10006 m6 on m6.cidconce01 = m8.cidconce01" +
" join mgw10007 m7 on m8.ciddocum02 = m7.ciddocum01 and m6.ciddocum01 = m7.ciddocum01 " +
" join mgw10009 m9 on m9.ciddocum02 = m8.ciddocum01 " +
" join mgw10008 m8a on m9.ciddocum01 = m8a.ciddocum01 " +
" join mgw10006 m6a on m6a.cidconce01 = m8a.cidconce01 " +
" where m8.ccancelado = 0 and m7.cnatural01 =0 " +
" and dtos(m9.cfechaab01) between '" + sfecha1 + "' and '" + sfecha2 + "' and m8a.ccancelado = 0   " +
" and m8.cimpuesto2 > 0 " +
" order by m9.cfechaab01, m8a.cfolio";


            x.mTraerInformacionPrimerReporte(lquery, comboBox1.SelectedValue.ToString());

            string lfechai = dateTimePicker1.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Year.ToString().PadLeft(4, '0');
            string lfechaf = dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Year.ToString().PadLeft(4, '0');


            x.mReporteIEPS(comboBox1.Text, lfechai, lfechaf);
            /*SELECT m8.cfolio as foliocargo, m8.cseriedo01 as seriecargo, m8.cfecha as fechacargo, m2.crazonso01 as cliente, m6.cnombrec01 as conceptocargo, ;
m8.cneto as netocargo, m8.cimpuesto1 as iva, m8.cimpuesto2 as ieps, m8.ctotal as totalcargo, ;
m8a.cfolio as folioabono, m8a.cseriedo01 as serieabono, m8a.cfecha as fechaabono, m6a.cnombrec01 as conceptoabono, m8a.ctotal as totalabono;
FROM MGW10008 m8 join mgw10002 m2 ;
on m8.cidclien01 = m2.cidclien01 ;
join mgw10006 m6 on m6.cidconce01 = m8.cidconce01;
join mgw10007 m7 on m8.ciddocum02 = m7.ciddocum01 and m6.ciddocum01 = m7.ciddocum01;
join mgw10009 m9 on m9.ciddocum02 = m8.ciddocum01;
join mgw10008 m8a on m9.ciddocum01 = m8a.ciddocum01;
join mgw10006 m6a on m6a.cidconce01 = m8a.cidconce01;
where m8.ccancelado = 0 and m7.cnatural01 =0;*/
        }
    }
}
