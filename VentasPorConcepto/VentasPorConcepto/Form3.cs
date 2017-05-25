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
    public partial class Form3 : Form
    {
        Class1 x = new Class1();
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            
            this.Text = " REPORTE PEDIDOS/FACTURAS";
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
            lquery = " select substr(dtos(ped.cfecha),7,2) + '/' + substr(dtos(ped.cfecha),5,2) + '/' + substr(dtos(ped.cfecha),1,4), ped.cseriedo01, ped.cfolio, cliente.crazonso01, max(ped.cneto), sum(movped.cunidade03*movped.cprecioc01) as pendiente, fac.cseriedo01,fac.cfolio, cliente.crazonso01, max(fac.cneto), max(fac.cimpuesto1), max(fac.cretenci02), max(fac.ctotal), max(movped.cunidade03),max(movped.cprecioc01), ped.ciddocum01 " +
" from mgw10008 ped  " +
" join mgw10002 cliente on cliente.cidclien01 = ped.cidclien01  " +
" join mgw10010 movped on ped.ciddocum01 = movped.ciddocum01  " +
" left join mgw10010 movfac on movped.cidmovim01 = movfac.cidmovto02  " +
" left join mgw10008 fac on movfac.ciddocum01 = fac.ciddocum01  " +
" where ped.ciddocum02 = 2  " +
" and fac.ciddocum02 = 4  " +
            " and dtos(ped.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "'" +
            " and ped.ccancelado = 0 " +
            " group by ped.ciddocum01 , ped.cfecha, ped.cseriedo01, ped.cfolio, cliente.crazonso01, fac.cseriedo01,fac.cfolio   ";
            lquery += " order by ped.cfecha, ped.cfolio";


            lquery = "select substr(dtos(ped.cfecha),7,2) + '/' + substr(dtos(ped.cfecha),5,2) + '/' + substr(dtos(ped.cfecha),1,4),ped.cseriedo01, ped.cfolio, cliente.crazonso01,avg(ped.cneto), sum(movped.cprecioc01*val(transform(movped.cunidade03,'9999.99'))) as pendiente,  " +
" fac.cseriedo01,fac.cfolio, cliente.crazonso01,avg(fac.cneto), avg(fac.cimpuesto1), avg(fac.cretenci02), avg(fac.ctotal), substr(dtos(fac.cfecha),7,2) + '/' + substr(dtos(fac.cfecha),5,2) + '/' + substr(dtos(fac.cfecha),1,4)  " +
 " from mgw10010 movped  " +
" join mgw10008 ped on movped.ciddocum01 = ped.ciddocum01  " +
" join mgw10002 cliente on cliente.cidclien01 = ped.cidclien01    " +
" left join mgw10010 m10f on m10f.cidmovto02 = movped.cidmovim01  " +
" left join mgw10008 fac on fac.ciddocum01 = m10f.ciddocum01 " +
" where  movped.ciddocum02 = 2  " +
" and dtos(ped.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' and ped.ccancelado = 0   " +
" group by  ped.cfecha, ped.cseriedo01, ped.cfolio, cliente.crazonso01, fac.cseriedo01,fac.cfolio, fac.cfecha     " +
" order by ped.cfecha, ped.cfolio";


            x.mTraerInformacionPrimerReporte(lquery, comboBox1.SelectedValue.ToString());

            string lfechai = dateTimePicker1.Value.Day.ToString().PadLeft(2,'0') + '/' + dateTimePicker1.Value.Month.ToString().PadLeft(2,'0') + '/' + dateTimePicker1.Value.Year.ToString().PadLeft(4,'0');
            string lfechaf = dateTimePicker2.Value.Day.ToString().PadLeft(2,'0') + '/' + dateTimePicker2.Value.Month.ToString().PadLeft(2,'0') + '/' + dateTimePicker2.Value.Year.ToString().PadLeft(4,'0');

                
            x.mReportePedidoFactura(comboBox1.Text, lfechai, lfechaf);
        }
    }
}
