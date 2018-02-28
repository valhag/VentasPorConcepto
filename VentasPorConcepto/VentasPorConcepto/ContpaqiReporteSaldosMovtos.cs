using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace VentasPorConcepto
{
    using MyExcel = Microsoft.Office.Interop.Excel;
    public partial class ContpaqiReporteSaldosMovtos : Form
    {
        

        string Cadenaconexion = "data source =" + Properties.Settings.Default.server +
        ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
        "; password = " + Properties.Settings.Default.password + ";";

        public ContpaqiReporteSaldosMovtos()
        {
            InitializeComponent();
        }

        private void ContpaqiSaldosMovtos_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 2;
            comboBox3.SelectedIndex = 0;
            this.Text = "Generacion txt desde ContPAQ i Contabilidad";

            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = Properties.Settings.Default.database;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;

            Properties.Settings.Default.Save();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (mValida())
            {
                
            Properties.Settings.Default.server = txtServer.Text;
            Properties.Settings.Default.database = txtBD.Text;
            Properties.Settings.Default.user = txtUser.Text;
            Properties.Settings.Default.password = txtPass.Text;

            Properties.Settings.Default.Save();
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
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

        private void button1_Click(object sender, EventArgs e)
        {
            DataSet info = new DataSet();
            mTraerInfo(ref info);
            mExcel(info);
            MessageBox.Show("Proceso Terminado");

        }

        private MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            return newWorkbook;

        }

        private string formateacuenta(string acuenta, List<int> x)
        {

            string lcuentaformateada = "";
            int indice = 0;
            foreach (int j in x)
            {
                lcuentaformateada += acuenta.Substring(indice, j);
                lcuentaformateada += "-";
                indice += j;
            }
            return lcuentaformateada.Substring(0, lcuentaformateada.Length - 1);
            
        }

        private void mExcel(DataSet info)
        {
            DataTable saldos = info.Tables[0];
            DataTable parametros = info.Tables[1];
            //DataTable abonos = info.Tables[2];

            string lmascarilla = parametros.Rows[0][0].ToString();
            List<int> x = new List<int>();
            var charsToRemove = new string[] { "-" };

            for (int i = 0; i < lmascarilla.Length; i++)
            {
                if (lmascarilla[i] != '-')
                {
                    x.Add(int.Parse(lmascarilla[i].ToString()));
                }
            }

            

            int lperiodo = comboBox1.SelectedIndex;
            string larchivo = textBox1.Text;
            string linea = "";
            string lcuenta = "1105002003000";
            
            
            //string[] lines={};

            using
                (System.IO.StreamWriter file =
                    new System.IO.StreamWriter(larchivo))
            {
                file.WriteLine(comboBox2.SelectedItem.ToString());
                file.WriteLine(lperiodo+1.ToString());
                foreach (DataRow row in saldos.Rows)
                {
                    if (row["codigo"].ToString() != "_ESTADISTICAS" && row["codigo"].ToString() != "_CUADRE" && row["codigo"].ToString() != "_ORDEN")
                    {
                        lcuenta = formateacuenta(row["codigo"].ToString(),x );
                        decimal acumulado=0;
                        decimal enperiodo = 0;
                        // checar periodo, 

                        acumulado = decimal.Parse(row[lperiodo + 6].ToString());
                        acumulado = decimal.Round(acumulado, 2);
                        enperiodo = acumulado - (decimal.Parse(row[lperiodo + 5].ToString()));
                        enperiodo = decimal.Round(enperiodo, 2);
                        /*enperiodo = decimal.Parse(row[lperiodo + 6].ToString());

                        for (int xx = 5; xx < lperiodo + 6; xx++)
                        {
                            acumulado += decimal.Parse(row[xx].ToString());
                        }*/



                        linea = "4001,"+ lcuenta + "," + enperiodo + "," + acumulado.ToString() ;
                        file.WriteLine(linea);
                    }
                }
                file.Close();


            }
            
            //lines[0] = comboBox2.SelectedItem.ToString();
            //lines[1] = lperiodo.ToString();

            //int indice = 2;
            //string linea = "";
            //foreach (DataRow row in saldos.Rows)
            //    lines[indice++] = "0";
            



            //string lcuenta1 = "";
            //string lcuenta2 = "";
            //string lcuenta3 = "";
            //string lcuenta4 = "";
            
            ////checar el combobox de filtro para ver si es por saldos
            //// si es por salods
            //int lindice = -1;
            //if (comboBox3.SelectedItem.ToString() == "Saldos No Cero")
            //{
            //    foreach (DataRow row in saldos.Rows)
            //    {
            //        lindice++;
            //        if (lcuenta1 != row[1].ToString().Trim())
            //        {
            //            if (decimal.Parse(row[4].ToString().Trim()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[1].ToString().Trim();
            //                sheet.Cells[lrenglon, 4].value = row[2].ToString().Trim();
            //                sheet.Cells[lrenglon, 5].value = row[3].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][2].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][2].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[4].ToString().Trim();
            //                lcuenta1 = row[1].ToString().Trim();
            //                lrenglon++;
            //            }
            //        }

            //        if (lcuenta2 != row[6].ToString().Trim())
            //        {
            //            if (decimal.Parse(row[9].ToString().Trim()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[6].ToString().Trim(); // cuenta
            //                sheet.Cells[lrenglon, 4].value = row[7].ToString().Trim(); // nombre cuenta
            //                sheet.Cells[lrenglon, 5].value = row[8].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][5].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][5].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[9].ToString().Trim();
            //                lcuenta2 = row[6].ToString().Trim();

            //                lrenglon++;
            //            }
            //        }
            //        if (lcuenta3 != row[11].ToString().Trim())
            //        {
            //            if (decimal.Parse(row[14].ToString().Trim()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[11].ToString().Trim(); // cuenta
            //                sheet.Cells[lrenglon, 4].value = row[12].ToString().Trim(); // nombre cuenta
            //                sheet.Cells[lrenglon, 5].value = row[13].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][8].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][8].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[14].ToString().Trim();
            //                lcuenta3 = row[11].ToString().Trim();
            //                lrenglon++;
            //            }
            //        }

            //        if (int.Parse(row[15].ToString()) != 0)
            //        {
            //            if (lcuenta4 != row[16].ToString().Trim())
            //            {
            //                if (decimal.Parse(row[19].ToString().Trim()) != 0)
            //                {
            //                    sheet.Cells[lrenglon, 3].value = "'" + row[16].ToString().Trim(); // cuenta
            //                    sheet.Cells[lrenglon, 4].value = row[17].ToString().Trim(); // nombre cuenta
            //                    sheet.Cells[lrenglon, 5].value = row[18].ToString().Trim();
            //                    sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][11].ToString();
            //                    sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][11].ToString();
            //                    sheet.Cells[lrenglon, 8].value = row[19].ToString().Trim();
            //                    lcuenta4 = row[16].ToString().Trim();
            //                    lrenglon++;
            //                }
            //            }
            //            if (int.Parse(row[20].ToString()) != 0)
            //                if (decimal.Parse(row[24].ToString().Trim()) != 0)
            //                {
            //                    sheet.Cells[lrenglon, 3].value = "'" + row[21].ToString().Trim(); // cuenta
            //                    sheet.Cells[lrenglon, 4].value = row[22].ToString().Trim(); // nombre cuenta
            //                    sheet.Cells[lrenglon, 5].value = row[23].ToString().Trim();
            //                    sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][14].ToString();
            //                    sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][14].ToString();
            //                    sheet.Cells[lrenglon, 8].value = row[24].ToString().Trim();

            //                    lrenglon++;
            //                }
            //        }

            //    }
            //}
            //else
            //{
            //    foreach (DataRow row in saldos.Rows)
            //    {
            //        lindice++;
            //        if (lcuenta1 != row[1].ToString().Trim())
            //        {
            //            if (decimal.Parse(cargos.Rows[lindice][2].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][2].ToString()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[1].ToString().Trim();
            //                sheet.Cells[lrenglon, 4].value = row[2].ToString().Trim();
            //                sheet.Cells[lrenglon, 5].value = row[3].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][2].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][2].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[4].ToString().Trim();
            //                lcuenta1 = row[1].ToString().Trim();
            //                lrenglon++;
            //            }
            //        }

            //        if (lcuenta2 != row[6].ToString().Trim())
            //        {
            //            if (decimal.Parse(cargos.Rows[lindice][5].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][5].ToString()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[6].ToString().Trim(); // cuenta
            //                sheet.Cells[lrenglon, 4].value = row[7].ToString().Trim(); // nombre cuenta
            //                sheet.Cells[lrenglon, 5].value = row[8].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][5].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][5].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[9].ToString().Trim();
            //                lcuenta2 = row[6].ToString().Trim();

            //                lrenglon++;
            //            }
            //        }
            //        if (lcuenta3 != row[11].ToString().Trim())
            //        {
            //            if (decimal.Parse(cargos.Rows[lindice][8].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][8].ToString()) != 0)
            //            {
            //                sheet.Cells[lrenglon, 3].value = "'" + row[11].ToString().Trim(); // cuenta
            //                sheet.Cells[lrenglon, 4].value = row[12].ToString().Trim(); // nombre cuenta
            //                sheet.Cells[lrenglon, 5].value = row[13].ToString().Trim();
            //                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][8].ToString();
            //                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][8].ToString();
            //                sheet.Cells[lrenglon, 8].value = row[14].ToString().Trim();
            //                lcuenta3 = row[11].ToString().Trim();
            //                lrenglon++;
            //            }
            //        }

            //        if (int.Parse(row[15].ToString()) != 0)
            //        {
            //            if (lcuenta4 != row[16].ToString().Trim())
            //            {
            //                if (decimal.Parse(cargos.Rows[lindice][11].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][11].ToString()) != 0)
            //                {
            //                    sheet.Cells[lrenglon, 3].value = "'" + row[16].ToString().Trim(); // cuenta
            //                    sheet.Cells[lrenglon, 4].value = row[17].ToString().Trim(); // nombre cuenta
            //                    sheet.Cells[lrenglon, 5].value = row[18].ToString().Trim();
            //                    sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][11].ToString();
            //                    sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][11].ToString();
            //                    sheet.Cells[lrenglon, 8].value = row[19].ToString().Trim();
            //                    lcuenta4 = row[16].ToString().Trim();
            //                    lrenglon++;
            //                }
            //            }
            //            if (int.Parse(row[20].ToString()) != 0)
            //                if (decimal.Parse(cargos.Rows[lindice][14].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][14].ToString()) != 0)
            //                {
            //                    sheet.Cells[lrenglon, 3].value = "'" + row[21].ToString().Trim(); // cuenta
            //                    sheet.Cells[lrenglon, 4].value = row[22].ToString().Trim(); // nombre cuenta
            //                    sheet.Cells[lrenglon, 5].value = row[23].ToString().Trim();
            //                    sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][14].ToString();
            //                    sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][14].ToString();
            //                    sheet.Cells[lrenglon, 8].value = row[24].ToString().Trim();

            //                    lrenglon++;
            //                }
            //        }

            //    }

            //}
            //sheet.Cells.EntireColumn.AutoFit();


        }

        private void mTraerInfo(ref DataSet info)
        {
            int aindicesaldomenor=0;
            int aindicesaldomayor=0;
            aindicesaldomayor = comboBox1.SelectedIndex + 1;

            int aejercicio = comboBox2.SelectedIndex;
            string lejercicio = comboBox2.SelectedItem.ToString();
            string lcampoinicial = "Importes";


            if (aindicesaldomayor == 1)
            {
                lcampoinicial = "SaldoIni";
                aindicesaldomayor = 1;
                aindicesaldomenor = 0;
            }
            else
                aindicesaldomenor = aindicesaldomayor - 1;

            string lindicesaldomenor = aindicesaldomenor.ToString();
            if (lindicesaldomenor == "0")
                lindicesaldomenor = "";
            string lindicesaldomayor = aindicesaldomayor.ToString();


            

            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);
            SqlCommand mySqlCommand = new SqlCommand();
            string saldos = "select '4001' fijo, c.codigo, c.id, c.nombre, e.Ejercicio, s.SaldoIni, s.Importes1, s.Importes2,s.Importes3,s.Importes4,s.Importes5,s.Importes6,s.Importes7,s.Importes8,s.Importes9,s.Importes10,s.Importes11,s.Importes12,s.Importes13,s.Importes14 " +
            " from cuentas c join saldoscuentas s on c.Id = s.IdCuenta and s.Tipo = 1 " +
            " join ejercicios e on e.Id = s.Ejercicio and e.Ejercicio = " + lejercicio +
            " where Afectable = 1 " +
            " or (c.tipo ='I' or c.tipo ='J') order by codigo";


            /*saldos = "select '4001' fijo, c.codigo, c.id, c.nombre, e.Ejercicio, s.SaldoIni, s.Importes1, s.Importes2,s.Importes3,s.Importes4,s.Importes5,s.Importes6, " +
" s.Importes7,s.Importes8,s.Importes9,s.Importes10,s.Importes11,s.Importes12,s.Importes13,s.Importes14  " +
" from cuentas c join saldoscuentas s on c.Id = s.IdCuenta and s.Tipo = 1  join ejercicios e on e.Id = s.Ejercicio and e.Ejercicio = 2015 " +
" where codigo = '1113010000'";*/

            saldos += " select EstructCta, mascarilla from Parametros";

 

            mySqlCommand.CommandText = saldos;

            DataSet ds = new DataSet();
            mySqlCommand.CommandType = CommandType.Text;
            mySqlCommand.Connection = DbConnection;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);
            info = ds;


        }
    }
}
