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
    public partial class NosSerieAmco : Form
    {
        ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        Class1 x = new Class1();
        public NosSerieAmco()
        {
            InitializeComponent();
        }

        private void NosSerieAmco_Load(object sender, EventArgs e)
        {
            this.Text = " Reporte Numero de Serie " + " " + this.ProductVersion;
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
        }

        private void mProcesar()
        {
            StringBuilder lquery = new StringBuilder();



            lquery.Append("SELECT p.CCODIGOPRODUCTO,a.CNOMBREALMACEN,max(ms.CFECHA), p.cnombreproducto, a.ccodigoalmacen from admNumerosSerie s ");
            lquery.Append("join admMovimientosSerie ms on s.CIDSERIE = ms.CIDSERIE ");
            lquery.Append("join admMovimientos m on ms.CIDMOVIMIENTO = m.CIDMOVIMIENTO ");
            lquery.Append("JOIN admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("JOIN admAlmacenes a on a.CIDALMACEN= s.CIDALMACEN ");
            lquery.Append("join admDocumentos d on d.CIDDOCUMENTO = m.CIDDOCUMENTO ");
            lquery.Append("where CNUMEROSERIE = '" + textBox1.Text + "' ");
            lquery.Append("group by p.CCODIGOPRODUCTO,a.CNOMBREALMACEN,p.cnombreproducto, a.ccodigoalmacen  ");

            //lquery.Append(" and dtos(d.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' and d.ccancelado = 0 " );





            //lquery.Append(" order by m8.cfecha, m8.cfolio ");

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);
            if (x.DatosReporte.Rows.Count > 0)
            {
                txtCodigoProducto.Text = x.DatosReporte.Rows[0][0].ToString();
                txtNombreProducto.Text = x.DatosReporte.Rows[0][3].ToString();
                txtCodigoAlmacen.Text = x.DatosReporte.Rows[0][4].ToString();
                txtNombreAlmacen.Text = x.DatosReporte.Rows[0][1].ToString();
                textBox4.Text = x.DatosReporte.Rows[0][2].ToString().Substring(0, 10);

                Clipboard.Clear();
                try
                {
                    //Clipboard.SetText("Serie:\t " + textBox1.Text + "\t" + " Codigo Producto: " + "\t" + txtCodigoProducto.Text.Trim() + "\tNombre de Producto\t" + txtNombreProducto.Text.Trim() + "\t" + " Codigo Almacen: " + "\t" + txtCodigoAlmacen.Text.Trim() + "\t" + "Nombre Almacen:\t" + txtNombreAlmacen.Text +"\t Fecha: " + "\t" + textBox4.Text);
                    Clipboard.SetText(textBox1.Text);

                }
                catch (Exception eee)
                {
                    //Clipboard.SetText("Serie:\t " + textBox1.Text + "\t" + " Codigo Producto: " + "\t" + txtCodigoProducto.Text.Trim() + "\tNombre de Producto\t" + txtNombreProducto.Text.Trim() + "\t" + " Codigo Almacen: " + "\t" + txtCodigoAlmacen.Text.Trim() + "\t" + "Nombre Almacen:\t" + txtNombreAlmacen.Text + "\t Fecha: " + "\t" + textBox4.Text);
                    Clipboard.SetText(textBox1.Text);
                }
            }


            //x.mTraerInformacionPedidoFactura(lquery, empresasComercial1.micombo.
            //  comboBox1.SelectedValue.ToString());

            //x.mReportePedidoFacturaComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, dateTimePicker2.Value);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            mProcesar();
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            mProcesar();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
