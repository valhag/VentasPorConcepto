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
using System.Data.SqlClient;

namespace VentasPorConcepto
{
    public partial class VentasSegmentoNegocio : ReporteBase
    {

        string CadenaConexionC = "";
        public VentasSegmentoNegocio()
        {
            
            InitializeComponent();
            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
            empresasComercial2.SelectedItem += new EventHandler(OnComboChangeC);
            //codigocatalogocomercial1.TextLeave+= new EventHandler(OnTextLeave1);
            //codigocatalogocomercial2.TextLeave += new EventHandler(OnTextLeave2);
            codigocatalogocomercial1.mSeteartipo(2, 0);
            codigocatalogocomercial2.mSeteartipo(2, 0);
            codigocatalogocomercial1.mSetLabelText("Cliente Inicial");
            codigocatalogocomercial2.mSetLabelText("Cliente Final");
            codigocatalogocomercial3.mSeteartipo(4, 0);
            codigocatalogocomercial4.mSeteartipo(4,0);
            codigocatalogocomercial3.mSetLabelText("Producto Inicial");
            codigocatalogocomercial4.mSetLabelText("Producto Final");


            
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

            
            //   dateTimePicker1.Value = DateTime.Parse("01/01/2014");
           //  dateTimePicker2.Value = DateTime.Parse("31/12/2014"); ;

        }

        private void OnComboChangeC(object sender, EventArgs e)
        {
            string x = empresasComercial2.aliasbdd;
            string y = empresasComercial2.Name;
            Properties.Settings.Default.RutaEmpresaC = empresasComercial2.aliasbdd;
            Properties.Settings.Default.Save();

        }

        private bool mValida()
        {
            string Cadenaconexion = "data source =" + txtServer.Text + ";initial catalog =" + txtBD.Text + ";user id = " + txtUser.Text + "; password = " + txtPass.Text + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        private void VentasSegmentoNegocio_Load(object sender, EventArgs e)
        {

            txtServer.Text = Properties.Settings.Default.server2;
            txtBD.Text = Properties.Settings.Default.database2;
            txtUser.Text = Properties.Settings.Default.user2;
            txtPass.Text = Properties.Settings.Default.password2;

            Properties.Settings.Default.Save();


            this.Text = " Reporte Venta Segmento de Negocio " + " " + this.ProductVersion;
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());


            //this.codigocatalogocomercial1.mSetLibreria(lrn, Cadenaconexion);

            // this.codigocatalogocomercial1.mSetLibreria(lrn);
            // this.codigocatalogocomercial1.mSeteartipo(1, 1);



            //Form5 xx = new Form5();
            //xx.ShowDialog();

            string server = Properties.Settings.Default.server;
            //MessageBox.Show("server " + server);
            if (Properties.Settings.Default.server != "")
            {

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                CadenaConexionC = "data source =" + Properties.Settings.Default.server2 +
                ";initial catalog =" + Properties.Settings.Default.database2 + " ;user id = " + Properties.Settings.Default.user2 +
                "; password = " + Properties.Settings.Default.password2 + ";";
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

                empresasComercial2.PopulateC(CadenaConexionC);


                //this.codigocatalogocomercial1.mSetConexion(x);
            }
            else
            {
                Form4 x = new Form4();
                x.Show();
            }
            empresasComercial2.SetTitulo("DATOS CONTABILIDAD");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (mValida())
            {
                Properties.Settings.Default.server2 = txtServer.Text;
                Properties.Settings.Default.database2 = txtBD.Text;
                Properties.Settings.Default.user2 = txtUser.Text;
                Properties.Settings.Default.password2 = txtPass.Text;

                Properties.Settings.Default.Save();
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");

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

            StringBuilder lquery = new StringBuilder();


            lquery.Append("select *, "); 
            lquery.Append("case when cestado = 0 and ccancelado = 0 then ' Sin Timbrar' ");
            lquery.Append("when ccancelado = 1 then 'Cancelado' ");
            lquery.Append("else ");
            lquery.Append("                'Activo' end cestado2 ");
            lquery.Append("from ");
            lquery.Append("( ");
            lquery.Append("select c.CNOMBRECONCEPTO ");
lquery.Append(", d.cfecha ");
            lquery.Append(", d.cfolio ");
            lquery.Append(", cl.CRAZONSOCIAL ");
            lquery.Append(", p.CCODIGOPRODUCTO ");
            lquery.Append(", p.CNOMBREPRODUCTO ");
            lquery.Append(", m.CUNIDADES ");
            lquery.Append(", m.cneto ");
            lquery.Append(", m.CTOTAL ");
            lquery.Append(", m.CIMPUESTO1 ");
            lquery.Append(", mo.CNOMBREMONEDA ");
            lquery.Append(", d.Ctipocambio ");
            lquery.Append(", m.CSCMOVTO ");
            lquery.Append(", d.ccancelado ");
            lquery.Append(", case when fd.cestado is null then 1 else fd.cestado end cestado ");
            lquery.Append(" from ");
            lquery.Append(" admdocumentos d ");
            lquery.Append("join admConceptos c on d.CIDCONCEPTODOCUMENTO = c.CIDCONCEPTODOCUMENTO ");
            lquery.Append("join admClientes cl on cl.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR ");
            lquery.Append("join admMovimientos m on m.CIDDOCUMENTO = d.CIDDOCUMENTO ");
            lquery.Append("join admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("join admMonedas mo on mo.CIDMONEDA = d.CIDMONEDA ");
            lquery.Append("left join admFoliosDigitales fd on fd.ciddocto = d.ciddocumento ");
            lquery.Append("where d.CIDDOCUMENTODE = 4 ");
            

            lquery.Append(" and m.cfecha >= '" + sfecha1 + "' and m.cfecha <= '" + sfecha2 + "'");



            if (lidcliente1 > 0 || lidcliente2 > 0)
                if (lidcliente1 > 0)
                    lquery.Append(" and d.CIDCLIENTEPROVEEDOR >= " + lidcliente1);
            if (lidcliente2 > 0)
                lquery.Append(" and d.CIDCLIENTEPROVEEDOR <= " + lidcliente2);

            if (lidproducto1 > 0 || lidproducto2 > 0)
                if (lidproducto1 > 0)
                    lquery.Append(" and m.CIDPRODUCTO >= " + lidproducto1);
            if (lidproducto2 > 0)
                lquery.Append(" and m.CIDPRODUCTO <= " + lidproducto2);
            lquery.Append(") a");





            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);


            x.mReporteVentasSNComercial(empresasComercial1.aliasbdd);
        }
    }
}
