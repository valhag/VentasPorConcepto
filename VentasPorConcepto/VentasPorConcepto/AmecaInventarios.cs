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
    public partial class AmecaInventarios : VentasPorConcepto.ReporteBase
    {
        public AmecaInventarios()
        {

            InitializeComponent();
        }

        private void AmecaInventarios_Load(object sender, EventArgs e)
        {
            this.Text = " Reporte Ameca Inventarios ";

        }

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            //string lquery;

            // string archivo = @"C:\fromgithub\archivotest.xlsx";

            //string archivo = @textBox1.Text;

            StringBuilder lquery = new StringBuilder();

            lquery.Append("select c.CNOMBRECONCEPTO,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO " );
  lquery.Append("  ,sum(m.CUNIDADES) as cantidad, sum(m.CCOSTOESPECIFICO)/ sum(m.cunidades) as costo, sum(CCOSTOESPECIFICO) as total " );
  lquery.Append("   from admMovimientos m  " );
  lquery.Append("   join admDocumentosModelo dm	on m.CIDDOCUMENTODE = dm.CIDDOCUMENTODE " );
  lquery.Append(" join admAlmacenes a on a.CIDALMACEN = m.CIDALMACEN " );
  lquery.Append(" join admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO " );
  lquery.Append(" join admDocumentos d on d.CIDDOCUMENTO = m.CIDDOCUMENTO " );
  lquery.Append(" join admconceptos c on c.CIDCONCEPTODOCUMENTO = d.CIDCONCEPTODOCUMENTO " );
  lquery.Append(" where dm.CAFECTAEXISTENCIA in (1,2) " );
  lquery.Append(" and a.CCODIGOALMACEN  in ('1','2','3','4','9', '99') " );
  lquery.Append(" and m.CAFECTADOINVENTARIO = 1 " );
  lquery.Append(" and m.CIDDOCUMENTO <> 0 " );
  lquery.Append(" and d.cfecha between '" + sfecha1 + "' and '" + sfecha2 + "' ");
  lquery.Append(" group by c.CNOMBRECONCEPTO,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO " );
  lquery.Append(" order by c.CNOMBRECONCEPTO,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO " );

            /*lquery.Append(" and dtos(m8.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "'" );
lquery.Append(" order by m2.ccodigoc01");*/





            //x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);

            x.mTraerInformacionComercial(lquery,empresasComercial1.aliasbdd);


            //x.mReporteForrajera(seleccionEmpresa1.lnombreempresa, archivo);
            x.mReporteForrajeraComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, dateTimePicker2.Value);
        }
    }
}
