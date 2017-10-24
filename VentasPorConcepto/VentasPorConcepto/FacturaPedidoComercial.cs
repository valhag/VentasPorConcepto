using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using LibreriaDoctos;

namespace VentasPorConcepto
{
    public partial class FacturaPedidoComercial : Form
    {
        ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        Class1 x = new Class1();

        public FacturaPedidoComercial()
        {
            InitializeComponent();
        }


        private void FacturaPedidoComercial_Load(object sender, EventArgs e)
        {

            this.Text = " Reporte Pedido Factura " + " " + this.ProductVersion;
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());


            string server = Properties.Settings.Default.server;
            //MessageBox.Show("server " + server);
            if (Properties.Settings.Default.server != "")
            {

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                //Archivo = Properties.Settings.Default.archivo;
            }
            if (Cadenaconexion != "")
                empresasComercial1.Populate(Cadenaconexion);
            else
            {
                Form4 x = new Form4();
                x.Show();
            }

            //empresasComercial1.SelectedItem += new EventHandler(OnComboChange);


            /*
              select a.CCODIGOAGENTE, a.CNOMBREAGENTE, dp.cfolio, d.CFECHA,d.CFOLIO,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, 
c.cvalorclasificacion, mon.CNOMBREMONEDA, m.CPRECIOCAPTURADO, m.CUNIDADESCAPTURADAS, m.cneto, m.ctotal, d.CFECHAVENCIMIENTO
from admmovimientos m
join admProductos p  on m.CIDPRODUCTO = p.CIDPRODUCTO
join admDocumentos d on d.CIDDOCUMENTO = m.CIDDOCUMENTO
join admAgentes a on a.CIDAGENTE = d.CIDAGENTE
join admClasificacionesValores c on c.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION1
join admMonedas mon on mon.CIDMONEDA = d.CIDMONEDA
left join admMovimientos mp on mp.CIDMOVIMIENTO = m.CIDMOVTOORIGEN
and mp.CIDDOCUMENTODE = 2
left join admDocumentos dp on mp.CIDDOCUMENTO = dp.CIDDOCUMENTO
and c.CIDCLASIFICACION = 25
where m.CIDDOCUMENTODE = 4

select * from admClasificaciones order by CIDCLASIFICACION

select * from admClasificacionesValores

select * from admMonedas
              */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

           // string lquery;



            StringBuilder lquery = new StringBuilder();

            lquery.Append("select cl.ccodigocliente, cl.crazonsocial, a.CCODIGOAGENTE, a.CNOMBREAGENTE, dp.cfolio as foliopedido, d.CFECHA,d.CFOLIO as foliofactura,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, ");
            lquery.Append("c.cvalorclasificacion, mon.CNOMBREMONEDA, m.CPRECIOCAPTURADO, m.CUNIDADESCAPTURADAS, m.cneto, m.ctotal, d.CFECHAVENCIMIENTO, ");
            lquery.Append("c2.cvalorclasificacion as cvalorclasificacion2 , c3.cvalorclasificacion as cvalorclasificacion3, c4.cvalorclasificacion as cvalorclasificacion4,c5.cvalorclasificacion as cvalorclasificacion5,c6.cvalorclasificacion as cvalorclasificacion6, ");
            lquery.Append("m.cporcentajedescuento1, m.cdescuento1,m.Cimpuesto1 ");
            lquery.Append("from admmovimientos m ");
            lquery.Append("join admProductos p on m.CIDPRODUCTO = p.CIDPRODUCTO ");
            lquery.Append("join admDocumentos d on d.CIDDOCUMENTO = m.CIDDOCUMENTO ");
            lquery.Append("join admClientes cl on cl.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR ");
            lquery.Append("join admAgentes a on a.CIDAGENTE = d.CIDAGENTE ");
            lquery.Append("join admClasificacionesValores c on c.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION1 ");
            lquery.Append("join admClasificacionesValores c2 on c2.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION2 ");
            lquery.Append("join admClasificacionesValores c3 on c3.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION3 ");
            lquery.Append("join admClasificacionesValores c4 on c4.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION4 ");
            lquery.Append("join admClasificacionesValores c5 on c5.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION5 ");
            lquery.Append("join admClasificacionesValores c6 on c6.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION6 ");
            lquery.Append("join admMonedas mon on mon.CIDMONEDA = d.CIDMONEDA ");
            lquery.Append("left join admMovimientos mp on mp.CIDMOVIMIENTO = m.CIDMOVTOORIGEN ");
            lquery.Append("and mp.CIDDOCUMENTODE = 2 ");
            lquery.Append("left join admDocumentos dp on mp.CIDDOCUMENTO = dp.CIDDOCUMENTO ");
            lquery.Append("and (c.CIDCLASIFICACION = 25  or c2.CIDCLASIFICACION = 26 or c3.CIDCLASIFICACION = 27 or c4.CIDCLASIFICACION = 28 or c5.CIDCLASIFICACION = 29 or c5.CIDCLASIFICACION = 30)");
            lquery.Append("where m.CIDDOCUMENTODE = 4 ");
            lquery.Append(" and d.ccancelado = 0 ");
            lquery.Append(" and d.cfecha between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append(" order by d.cfecha asc, d.CFOLIO asc ");

//lquery.Append(" and dtos(d.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' and d.ccancelado = 0 " );





            //lquery.Append(" order by m8.cfecha, m8.cfolio ");

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);
            //x.mTraerInformacionPedidoFactura(lquery, empresasComercial1.micombo.
            //  comboBox1.SelectedValue.ToString());

            x.mReportePedidoFacturaComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, dateTimePicker2.Value);
        }
    }
}
