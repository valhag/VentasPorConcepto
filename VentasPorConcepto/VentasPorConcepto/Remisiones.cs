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
    public partial class Remisiones : Form
    {

        ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        Class1 x = new Class1();

        public Remisiones()
        {
            InitializeComponent();
        }

        private void Remisiones_Load(object sender, EventArgs e)
        {

            this.Text = " Reporte/Borrado Remisiones " + " " + this.ProductVersion;
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

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            // string lquery;



            StringBuilder lquery = new StringBuilder();

            lquery.Append("SELECT format(d.CFECHA,'dd/MM/yyyy') as FECHA, d.cfolio as FOLIO, c.CRAZONSOCIAL AS [RAZON SOCIAL],d.CTOTALUNIDADES AS [TOTAL UNIDADES], d.cneto AS NETO, d.CTOTAL AS TOTAL,co.ccodigoconcepto, d.cseriedocumento as SERIE ");
            lquery.Append("FROM admDocumentos d ");
            lquery.Append("JOIN admClientes c on c.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR ");
            lquery.Append("JOIN admConceptos co on co.CIDCONCEPTODOCUMENTO = d.CIDCONCEPTODOCUMENTO ");
            lquery.Append("where d.CCANCELADO = 0 and d.CIDDOCUMENTODE = 3 ");
            lquery.Append("and d.CFECHA between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append("order by Fecha ");

            lquery.Append("SELECT  p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, SUM(m.CUNIDADES) AS CUNIDADES, SUM(m.CNETO) AS CNETO, SUM(m.ctotal) AS CTOTAL ");
            lquery.Append("FROM admDocumentos d ");
            lquery.Append("JOIN admMovimientos m on d.CIDDOCUMENTO = m.CIDDOCUMENTO ");
            lquery.Append("join admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("where d.CCANCELADO = 0 and d.CIDDOCUMENTODE = 3 ");
            lquery.Append("and d.CFECHA between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append(" group by p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO ");
            lquery.Append("order by  p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO ;");
            
            

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);

            dataGridView1.DataSource = null;
            dataGridView1.DataSource = x.DatosReporte;
            dataGridView1.AutoResizeColumns();
            //dataGridView1.AutoGenerateColumns = false;
            /*DataGridViewCheckBoxColumn CBColumn = new DataGridViewCheckBoxColumn();
            CBColumn.HeaderText = "ColumnHeader";
            CBColumn.FalseValue = "0";
            CBColumn.TrueValue = "1";
            dataGridView1.Columns.Insert(0, CBColumn);*/

            dataGridView1.Columns[6].Visible = false;

            
            if (checkBox1.Checked == true)
                x.mReporteRemisionesComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, dateTimePicker2.Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            

            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();

            //lr.mborr
            RegConexion newcon = new RegConexion();
            newcon.database = empresasComercial1.aliasbdd;
            newcon.server = Properties.Settings.Default.server;
            newcon.usuario = Properties.Settings.Default.user;
            newcon.ps = Properties.Settings.Default.password;
            lrn.mAsignaEmpresaComercial(newcon);
            foreach (DataGridViewRow x in dataGridView1.Rows)
            {
                
                lrn.mBorrarDocto(x.Cells["ccodigoconcepto"].Value.ToString(), x.Cells["SERIE"].Value.ToString(), x.Cells["FOLIO"].Value.ToString());
            }
            MessageBox.Show("Proceso Terminado");
        }

        private void Remisiones_FormClosed(object sender, FormClosedEventArgs e)
        {
            lrn.mCerrarSdkComercial();
        }
    }
}
