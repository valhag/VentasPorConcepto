using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using MyExcel = Microsoft.Office.Interop.Excel;

namespace VentasPorConcepto
{
    public partial class ReporteNominas : Form
    {
        string Cadenaconexion;
        DataSet ds = new DataSet();
        public ReporteNominas()
        {
           
            InitializeComponent();
        }

        private void ReporteNominas_Load(object sender, EventArgs e)
        {
            Cadenaconexion = "";
            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = Properties.Settings.Default.database;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;
            if (txtServer.Text != "")
            {
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
        ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
        "; password = " + Properties.Settings.Default.password + ";";
                mLlenaPeriodos();
            }

            Properties.Settings.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            bool regresa = mValida();
            if (regresa == true)
            {
                MessageBox.Show("Parametros de conexion correctamente");
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;
                Properties.Settings.Default.Save();
                mLlenaPeriodos();
            }
            else
                MessageBox.Show("Configure parametros de conexion correctamente");
        }

        private void mLlenaPeriodos()
        {
            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);
            SqlCommand mySqlCommand = new SqlCommand();

            if (txtEjercicio.Text == "")
                return;

            string periodos = "select idperiodo," +
" CONVERT(nvarchar(30), fechainicio, 103) + '-'  +  CONVERT(nvarchar(30), fechafin, 103) as fecha" +
" from " +
" nom10002 " +
" where ejercicio = " + txtEjercicio.Text + " order by fechainicio";



            DataSet ds = new DataSet();

            //DbConnection = new SqlConnection(Cadenaconexion);
            mySqlCommand.CommandText = periodos;
            mySqlCommand.CommandType = CommandType.Text;
            mySqlCommand.Connection = DbConnection;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            
            try
            {
                comboBox1.DataSource = null;
                comboBox1.Items.Clear();
                comboBox1.DataSource = ds.Tables[0];
                comboBox1.DisplayMember = "Fecha";
                comboBox1.ValueMember = "idperiodo";
            }
            catch (Exception eee)
            {

            }

            //info = ds;

        }


        private void mTraerInfoNominas()
        {

            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);
            SqlCommand mySqlCommand = new SqlCommand();
            string saldos = "select e.codigoempleado, e.nombrelargo, e.curpi + substring(ltrim(str(year(fechanacimiento)) ),3,2) + right('00' + ltrim(str(month(fechanacimiento))),2) + right('00' + ltrim(str(day(fechanacimiento))),2) + e.homoclave as rfc, p.descripcion as puesto, d.descripcion as depto, " +
" c.tipoconcepto, c.numeroconcepto, c.descripcion as concepto,m.importetotal " + 
" , sum( case when c.tipoconcepto = 'P' then m.importetotal else 0 end) over(partition by e.idempleado) as x  " +
" , sum( case when c.tipoconcepto = 'D' then m.importetotal else 0 end) over(partition by e.idempleado) as y  " +
" , m.idperiodo  " +
" from nom10001 e  " +
" join nom10007 m  " +
" on e.idempleado = m.idempleado  " +
" join nom10006 p  " +
" on e.idpuesto = p.idpuesto  " +
" join nom10003 d  " +
" on e.iddepartamento = d.iddepartamento  " +
" join nom10004 c  " +
" on m.idconcepto = c.idconcepto  " +
" where c.tipoconcepto in ('P','D')  " +
" and idperiodo = " + comboBox1.SelectedValue.ToString();

            mySqlCommand.CommandText = saldos;

            
            mySqlCommand.CommandType = CommandType.Text;
            mySqlCommand.Connection = DbConnection;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            ds = null;
            ds = new DataSet();
            mySqlDataAdapter.Fill(ds);
            
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

        private void textBox1_Leave(object sender, EventArgs e)
        {
            mLlenaPeriodos();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count > 0)
            {
                if (comboBox1.SelectedIndex > -1)
                {
                    //MessageBox.Show(comboBox1.SelectedValue.ToString());
                    mTraerInfoNominas();
                    mReporteNominas();
                }
            }
        }
        public void mReporteNominas()
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            //int lrengloniniciaconcepto = 6;
            //int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezado(sheet, "Productos Pendientes de Surtir Ordenes de Compras", lrenglon);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            //lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;

                string sfecha = DateTime.Today.ToShortDateString();
                string dia = DateTime.Today.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Today.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Today.Year.ToString();
                sfecha = sfecha.Substring(6, 2).PadLeft(2, '0') + "/" + sfecha.Substring(4, 2).PadLeft(2, '0') + "/" + sfecha.Substring(0, 4);
                sheet.Cells[lrenglon, lcolumna++].value = dia;
                //sheet.Cells[lrenglon, lcolumna++].value = row["codigoempleado"].ToString().Trim(); 
                sheet.Cells[lrenglon, lcolumna++].value = row["nombrelargo"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["rfc"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["puesto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["depto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["x"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["y"].ToString().Trim();
                //sheet.Cells[lrenglon, lcolumna++].value = "0";
                sheet.Cells[lrenglon, lcolumna++].value = row["tipoconcepto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["numeroconcepto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["concepto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["importetotal"].ToString().Trim();
                lrenglon++;

                /*
                sheet.Cells[lrenglon, lcolumna++].value = "Fecha de pago";
                sheet.Cells[lrenglon, lcolumna++].value = "Nombre del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "RFC del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "Puesto del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "Area de trabajo";
                sheet.Cells[lrenglon, lcolumna++].value = "Total de percepciones";
                sheet.Cells[lrenglon, lcolumna++].value = "Total de deducciones";
                sheet.Cells[lrenglon, lcolumna++].value = "Total Liquido";
                sheet.Cells[lrenglon, lcolumna++].value = "Tipo de concepto (P=percepcion D=deduccion)";
                sheet.Cells[lrenglon, lcolumna++].value = "Clave del concepto";
                sheet.Cells[lrenglon, lcolumna++].value = "Concepto";
                sheet.Cells[lrenglon, lcolumna++].value = "Importe del concepto";*/
            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }
        public MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            return newWorkbook;

        }
        private void configuracionencabezado(MyExcel.Worksheet sheet, string texto, int lrenglon)
        {
            //EncabezadoEmpresa(sheet, "test", texto);
            int lcolumna = 1;



            /*sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";

            sheet.Cells[1, 7].value = lfecha1;
            sheet.Cells[2, 7].value = lfecha2;*/



                //            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


                sheet.Cells[lrenglon, lcolumna++].value = "Fecha de pago";
                sheet.Cells[lrenglon, lcolumna++].value = "Nombre del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "RFC del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "Puesto del empleado";
                sheet.Cells[lrenglon, lcolumna++].value = "Area de trabajo";
                sheet.Cells[lrenglon, lcolumna++].value = "Total de percepciones";
                sheet.Cells[lrenglon, lcolumna++].value = "Total de deducciones";
                //sheet.Cells[lrenglon, lcolumna++].value = "Total Liquido";
                sheet.Cells[lrenglon, lcolumna++].value = "Tipo de concepto (P=percepcion D=deduccion)";
                sheet.Cells[lrenglon, lcolumna++].value = "Clave del concepto";
                sheet.Cells[lrenglon, lcolumna++].value = "Concepto";
                sheet.Cells[lrenglon, lcolumna++].value = "Importe del concepto";
            
        }
        private void EncabezadoEmpresa(MyExcel.Worksheet sheet, string Empresa, string texto)
        {
            sheet.Cells[2, 3].value = Empresa;
            sheet.get_Range("C2").Font.Size = 16;
            sheet.get_Range("C2").Font.Bold = true;
            sheet.get_Range("C2", "D2").Merge();
            sheet.Cells[3, 3].value = texto;
            sheet.get_Range("C3", "E3").Merge();
            sheet.get_Range("C3").Font.Bold = true;
        }
        
        
    }
}
