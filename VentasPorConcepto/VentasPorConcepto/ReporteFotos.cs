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
    public partial class ReporteFotos : Form
    {
        Class1 x = new Class1();
        public ReporteFotos()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            string lquery;



            lquery = "select m2.ccodigoc01, m2.crazonso01, m8.creferen01, m8.ctextoex01, m8.ctextoex02, m8.ctextoex03, " + 
" m8.cfecha, m8.cfechaen01, m8.cfechave01, m5.cnomaltern, m5.ccodigop01, m5.cnombrep01, m10.cunidades, m26.cdesplie01, m10.cprecioc01, m10.ctotal,m10.creferen01 as creferen02, m10.ctextoex02, m10.ctextoex03, m10.cunidade03 " + 
" from mgw10008 m8 " + 
" join mgw10002 m2 on m8.cidclien01 = m2.cidclien01" + 
" join mgw10010 m10 on m10.ciddocum01 = m8.ciddocum01" + 
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01" + 
" join mgw10026 m26 on m5.cidunida01 = m26.cidunidad" + 
" where m8.ciddocum02 = 17 and m8.ccancelado = 0 " +
" and m10.cunidade03 > 0 " +
" and dtos(m8.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "'" +
" order by m2.ccodigoc01";





            x.mTraerInformacionPrimerReporte(lquery, seleccionEmpresa1.lrutaempresa);



            x.mReporteFotos(seleccionEmpresa1.lnombreempresa);
         //   x.mTestFotos();
        }

        private void ReporteFotos_Load(object sender, EventArgs e)
        {

        }
    }
}
