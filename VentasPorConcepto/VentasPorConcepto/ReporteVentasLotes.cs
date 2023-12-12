using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LibreriaDoctos;
using System.IO;

namespace VentasPorConcepto {
    public partial class ReporteVentasLotes : ReporteBase
    {
        protected ClassRN lrn = new ClassRN();

        public ReporteVentasLotes()
        {
            InitializeComponent();
            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
            //codigocatalogocomercial1.TextLeave+= new EventHandler(OnTextLeave1);
            //codigocatalogocomercial2.TextLeave += new EventHandler(OnTextLeave2);
            codigocatalogocomercial1.mSeteartipo(2, 0);
            codigocatalogocomercial2.mSeteartipo(2, 0);
            codigocatalogocomercial1.mSetLabelText("Cliente Inicial");
            codigocatalogocomercial2.mSetLabelText("Cliente Final");
            codigocatalogocomercial3.mSeteartipo(4, 0);
            codigocatalogocomercial4.mSeteartipo(4, 0);
            codigocatalogocomercial3.mSetLabelText("Producto Inicial");
            codigocatalogocomercial4.mSetLabelText("Producto Final");


            codigocatalogocomercial5.mSeteartipo(1, 0);
            codigocatalogocomercial6.mSeteartipo(1, 0);
            codigocatalogocomercial5.mSetLabelText("Agente Inicial");
            codigocatalogocomercial6.mSetLabelText("Agente Final");


        }

        private void OnComboChange(object sender, EventArgs e)
        {
            //MessageBox.Show("cambia");
            //lrn.almacenes = "chido";
            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            //Properties.Settings.Default.database = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();
            codigocatalogocomercial1.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";
            codigocatalogocomercial1.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial1.lrn = lrn;
            codigocatalogocomercial2.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial2.lrn = lrn;
            codigocatalogocomercial2.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";

            codigocatalogocomercial3.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial3.lrn = lrn;
            codigocatalogocomercial3.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";

            codigocatalogocomercial4.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial4.lrn = lrn;
            codigocatalogocomercial4.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";

            codigocatalogocomercial5.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial5.lrn = lrn;
            codigocatalogocomercial5.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";

            codigocatalogocomercial6.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial6.lrn = lrn;
            codigocatalogocomercial6.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";

         //   dateTimePicker1.Value = DateTime.Parse("01/01/2014");
           // dateTimePicker2.Value = DateTime.Parse("12/31/2014"); ;

        }

        private void ReporteVentasLotes_Load(object sender, EventArgs e)
        {

            this.Text = " Reporte Venta Lotes " + " " + this.ProductVersion;
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());


            //this.codigocatalogocomercial1.mSetLibreria(lrn, Cadenaconexion);

            // this.codigocatalogocomercial1.mSetLibreria(lrn);
            // this.codigocatalogocomercial1.mSeteartipo(1, 1);




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
            {
                empresasComercial1.Populate(Cadenaconexion);
                RegConexion x = new RegConexion();
                x.database = Properties.Settings.Default.database;
                x.server = Properties.Settings.Default.server;
                x.usuario = Properties.Settings.Default.user;
                x.ps = Properties.Settings.Default.password;
                //    lrn.lbd.miconexion._


                //this.codigocatalogocomercial1.mSetConexion(x);
            }
            else
            {
                Form4 x = new Form4();
                x.Show();
            }
        }

        private void codigocatalogocomercial6_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');


            long lidcliente1 = 0;
            long lidcliente2 = 0;
            if (codigocatalogocomercial1.lRegClienteProveedor != null)
                lidcliente1 = codigocatalogocomercial1.lRegClienteProveedor.Id;
            if (codigocatalogocomercial2.lRegClienteProveedor != null)
                lidcliente2 = codigocatalogocomercial2.lRegClienteProveedor.Id;

            long lidproducto1 = 0;
            long lidproducto2 = 0;
            if (codigocatalogocomercial4.lRegClienteProveedor != null)
                lidproducto1 = codigocatalogocomercial4.lRegClienteProveedor.Id;
            if (codigocatalogocomercial3.lRegClienteProveedor != null)
                lidproducto2 = codigocatalogocomercial3.lRegClienteProveedor.Id;

            long lidagente1 = 0;
            long lidagente2 = 0;
            if (codigocatalogocomercial6.lRegClienteProveedor != null)
                lidagente1 = codigocatalogocomercial6.lRegClienteProveedor.Id;
            if (codigocatalogocomercial5.lRegClienteProveedor != null)
                lidagente2 = codigocatalogocomercial5.lRegClienteProveedor.Id;

            //long lidcliente2 = codigocatalogocomercial2.lRegClienteProveedor.Id;
            //long lidalmacen2 = codigocatalogocomercial2.lRegClienteProveedor.Id;


            //string lquery;

            // string archivo = @"C:\fromgithub\archivotest.xlsx";
            StringBuilder lquery = new StringBuilder();



            lquery.Append("select  d.CSERIEDOCUMENTO ");
            lquery.Append(", d.CFOLIO ");
            lquery.Append(", c.CCODIGOCLIENTE");
            lquery.Append(", c.CRAZONSOCIAL");
            lquery.Append(", a.CNOMBREAGENTE");
            lquery.Append(", FORMAT (d.CFECHA, 'dd/MM/yyyy') as cfecha");
            lquery.Append(", d.CCANCELADO");
            lquery.Append(", m.cunidades");
            lquery.Append(", p.CCODIGOPRODUCTO");
            lquery.Append(", p.CNOMBREPRODUCTO");
            lquery.Append(", m.CPRECIO");
            lquery.Append(", m.cneto - (m.CDESCUENTO1 + m.CDESCUENTO2 + m.CDESCUENTO3 + m.CDESCUENTO4 + m.CDESCUENTO5) cneto");
            lquery.Append(", m.ctotal");
            lquery.Append(", co.CNOMBRECONCEPTO");
            lquery.Append(", d.CREFERENCIA");
            lquery.Append(", d.CTEXTOEXTRA1");
            lquery.Append(", d.COBSERVACIONES");
            lquery.Append(", d.CTOTALUNIDADES");
            lquery.Append(", ca.CNUMEROLOTE");
            lquery.Append(", ca.CEXISTENCIA");
            lquery.Append(", m.CIDMOVIMIENTO");
            lquery.Append(" from admDocumentos d");
            lquery.Append(" join admConceptos co on co.CIDCONCEPTODOCUMENTO = d.CIDCONCEPTODOCUMENTO");
            lquery.Append(" join admclientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR");
            lquery.Append(" join admAgentes a on d.CIDAGENTE = a.CIDAGENTE");
            lquery.Append(" join admMovimientos m on m.CIDDOCUMENTO = d.CIDDOCUMENTO");
            lquery.Append(" join admProductos p on m.CIDPRODUCTO = p.CIDPRODUCTO");
            lquery.Append(" join admMovimientosCapas mc on mc.CIDMOVIMIENTO = m.CIDMOVIMIENTO");
            lquery.Append(" join admCapasProducto ca on ca.CIDCAPA = mc.CIDCAPA");
            lquery.Append(" where d.CIDDOCUMENTODE = 4");

            lquery.Append(" and m.cfecha >= '" + sfecha1 + "' and m.cfecha <= '" + sfecha2 + "'");



            if (lidcliente1 > 0  || lidcliente2 > 0 )
                if (lidcliente1 > 0)
                    lquery.Append(" and d.CIDCLIENTEPROVEEDOR >= " + lidcliente1);
                if (lidcliente2 > 0)
                    lquery.Append(" and d.CIDCLIENTEPROVEEDOR <= " + lidcliente2);

            if (lidproducto1 > 0 || lidproducto2 > 0)
                if (lidproducto1 > 0)
                    lquery.Append(" and m.CIDPRODUCTO >= " + lidproducto1);
            if (lidproducto2 > 0)
                lquery.Append(" and m.CIDPRODUCTO <= " + lidproducto2);

            if (lidagente1 > 0 || lidagente2 > 0)
                if (lidagente1 > 0)
                    lquery.Append(" and d.CIDAGENTE >= " + lidagente1);
            if (lidagente2 > 0)
                lquery.Append(" and d.CIDAGENTE <= " + lidagente2);




            /*
                        lquery.Append(" and m.cfecha <= '" + sfecha1 + "' ");
                        lquery.Append(" and m.cidalmacen  between '" + lidalmacen1.ToString() + "' and '" + lidalmacen2.ToString() + "' ");
                        lquery.Append("group by p.cidproducto, a.cidalmacen, p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, a.CNOMBREALMACEN, m.CAFECTAEXISTENCIA, mm.CZONA ");
                        lquery.Append(") x ");
                        lquery.Append("group by x.cidproducto, x.cidalmacen, x.CCODIGOPRODUCTO, x.CNOMBREPRODUCTO, x.CNOMBREALMACEN, x.CZONA ");
                        lquery.Append(") y ");
                        lquery.Append("join tempcte ch on ch.CIDPRODUCTO = y.CIDPRODUCTO and ch.CIDALMACEN = y.CIDALMACEN ");
                        lquery.Append("and rownumber = 1 ");*/


            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);


            x.mReporteVentasLotesComercial(empresasComercial1.aliasbdd);

            MessageBox.Show("Reporte Terminado");
        }
    }
}
