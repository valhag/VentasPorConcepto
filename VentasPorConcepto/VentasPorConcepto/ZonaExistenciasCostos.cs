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


namespace VentasPorConcepto
{
    public partial class ZonaExistenciasCostos : ReporteBase
    {
        protected ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        Class1 x = new Class1();

        public ZonaExistenciasCostos()
        {
            InitializeComponent();
            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
            //codigocatalogocomercial1.TextLeave+= new EventHandler(OnTextLeave1);
            //codigocatalogocomercial2.TextLeave += new EventHandler(OnTextLeave2);
            codigocatalogocomercial1.mSeteartipo(5,0);
            codigocatalogocomercial2.mSeteartipo(5, 0);
            codigocatalogocomercial1.mSetLabelText("Almacen Inicial");
            codigocatalogocomercial2.mSetLabelText("Almacen Final");

            
        }

        public void OnTextLeave1(object sender, EventArgs e)
        {
            if (codigocatalogocomercial1.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(4, codigocatalogocomercial1.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial1.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Almacen no Existe");
                    codigocatalogocomercial1.mSetFocus();
                }
            }
        }

        public void OnTextLeave2(object sender, EventArgs e)
        {
            if (codigocatalogocomercial2.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(4, codigocatalogocomercial2.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial2.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Almacen no Existe");
                    codigocatalogocomercial2.mSetFocus();
                }
            }
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
            codigocatalogocomercial1.lrn.lbd.miconexion._NombreAplicacion = "VentasPorConcepto";
        }
        private void Form5_Load(object sender, EventArgs e)
        {
            this.Text = " Reportes Costos por Zona " + " " + this.ProductVersion;
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

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');


            long lidalmacen1 = codigocatalogocomercial1.lRegClienteProveedor.Id;
            long lidalmacen2 = codigocatalogocomercial2.lRegClienteProveedor.Id;


            //string lquery;

            // string archivo = @"C:\fromgithub\archivotest.xlsx";
            StringBuilder lquery = new StringBuilder();



            lquery.Append("WITH tempcte AS ");
lquery.Append("( ");
            lquery.Append("select *, ROW_NUMBER() over(partition by cidproducto order by cidcostoh desc) rownumber ");
            lquery.Append("from admCostosHistoricos ");
            lquery.Append("where CIDALMACEN between '" + lidalmacen1.ToString() + "' and '" + lidalmacen2.ToString() + "' ");
            lquery.Append("and CFECHACOSTOH <= '" + sfecha1 + "' ");
            lquery.Append(") ");

            lquery.Append("select y.CCODIGOPRODUCTO, y.CNOMBREPRODUCTO, y.CNOMBREALMACEN, y.cantidad, y.czona, ch.CCOSTOH costo from ");
            lquery.Append("( ");
            lquery.Append("select x.CCODIGOPRODUCTO, x.CNOMBREPRODUCTO, x.CNOMBREALMACEN, sum(unidades) cantidad, isNull(x.CZONA, '') CZONA, x.CIDPRODUCTO, x.CIDALMACEN ");
            lquery.Append("from( ");
            lquery.Append("select p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, a.CNOMBREALMACEN, p.CIDPRODUCTO, a.cidalmacen  ");
            lquery.Append(",case when m.CAFECTAEXISTENCIA = 1 then sum(m.cunidades) ");
            lquery.Append("when m.cafectaexistencia = 2 then sum(m.cunidades) * -1 end unidades, ");
            lquery.Append("mm.CZONA ");
            lquery.Append("from admMovimientos m ");
            lquery.Append("join admProductos p on m.cidproducto = p.cidproducto ");
            lquery.Append("join admAlmacenes a on a.CIDALMACEN = m.CIDALMACEN ");
            lquery.Append("left ");
            lquery.Append("join admMaximosMinimos mm on mm.CIDPRODUCTO = p.CIDPRODUCTO and a.CIDALMACEN = mm.CIDALMACEN ");
            lquery.Append("where CAFECTADOINVENTARIO = 1 ");
            //lquery.Append(" and CAFECTAEXISTENCIA = 1 ");
            lquery.Append(" and m.cfecha <= '" + sfecha1 + "' ");
            lquery.Append(" and m.cidalmacen  between '" + lidalmacen1.ToString() + "' and '" + lidalmacen2.ToString() + "' ");
            lquery.Append("group by p.cidproducto, a.cidalmacen, p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, a.CNOMBREALMACEN, m.CAFECTAEXISTENCIA, mm.CZONA ");
            lquery.Append(") x ");
            lquery.Append("group by x.cidproducto, x.cidalmacen, x.CCODIGOPRODUCTO, x.CNOMBREPRODUCTO, x.CNOMBREALMACEN, x.CZONA ");
            lquery.Append(") y ");
            lquery.Append("join tempcte ch on ch.CIDPRODUCTO = y.CIDPRODUCTO and ch.CIDALMACEN = y.CIDALMACEN ");
            lquery.Append("and rownumber = 1 ");

            /*            lquery.Append("select x.CCODIGOPRODUCTO, x.CNOMBREPRODUCTO, x.CNOMBREALMACEN, sum(unidades) cantidad, sum(costo) / 2 costo, isNull(x.CZONA, '') CZONA from ");
            lquery.Append("   (                                                                                                         ");
            lquery.Append("   select p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, a.CNOMBREALMACEN,                                            ");
            lquery.Append("   case when m.CAFECTAEXISTENCIA = 1 then sum(m.cunidades)                                                   ");
            lquery.Append("   when m.cafectaexistencia = 2 then sum(m.cunidades) * -1 end unidades,                                     ");
            lquery.Append("   case when m.CAFECTAEXISTENCIA = 1 then 'E'                                                                ");
            lquery.Append("when m.cafectaexistencia = 2 then  'S' end tipo                                                              ");
            lquery.Append(",case when m.CAFECTAEXISTENCIA = 1 then sum(m.CCOSTOESPECIFICO)                                              ");
            lquery.Append("when m.cafectaexistencia = 2 then sum(m.CCOSTOESPECIFICO) end costo                                          ");
            lquery.Append(", mm.CZONA                                                                                                   ");
            lquery.Append(" from admMovimientos m                                                                                       ");
            lquery.Append("join admProductos p on m.cidproducto = p.cidproducto                                                         ");
            lquery.Append("join admAlmacenes a on a.CIDALMACEN = m.CIDALMACEN                                                           ");
            lquery.Append("left join admMaximosMinimos mm on mm.CIDPRODUCTO = p.CIDPRODUCTO and a.CIDALMACEN = mm.CIDALMACEN            ");
            lquery.Append("where CAFECTADOINVENTARIO = 1                                                                                ");
                        lquery.Append(" and CAFECTAEXISTENCIA = 1 ");   
                        lquery.Append(" and m.cfecha <= '" + sfecha1 + "' ");

                        lquery.Append(" and m.cidalmacen  between '" + lidalmacen1.ToString() + "' and '" + lidalmacen2.ToString() + "' ");
                        lquery.Append(" group by p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, a.CNOMBREALMACEN, m.CAFECTAEXISTENCIA, mm.CZONA              ");
            lquery.Append(" ) x                                                                                                         ");
                        lquery.Append(" group by x.CCODIGOPRODUCTO, x.CNOMBREPRODUCTO, x.CNOMBREALMACEN, x.CZONA                                    ");
            lquery.Append(";");

                */
            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);


            x.mReporteZonaCostoComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, codigocatalogocomercial1.lRegClienteProveedor.RazonSocial, codigocatalogocomercial2.lRegClienteProveedor.RazonSocial);

        }
    }
}
