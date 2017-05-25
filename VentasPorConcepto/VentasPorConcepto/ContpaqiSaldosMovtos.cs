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
    public partial class ContpaqiSaldosMovtos : Form
    {
        

        string Cadenaconexion = "data source =" + Properties.Settings.Default.server +
        ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
        "; password = " + Properties.Settings.Default.password + ";";

        public ContpaqiSaldosMovtos()
        {
            InitializeComponent();
        }

        private void ContpaqiSaldosMovtos_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 2;
            comboBox3.SelectedIndex = 0;
            this.Text = "Reporte Contpaq i";

            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = Properties.Settings.Default.database;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password ;

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

        }

        private MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            return newWorkbook;

        }

        private void mExcel(DataSet info)
        {
            DataTable saldos = info.Tables[0];
            DataTable cargos = info.Tables[1];
            DataTable abonos = info.Tables[2];

            MyExcel.Workbook newWorkbook = mIniciarExcel();

            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];
            
                int lrenglon = 4;

            string lcuenta1 = "";
            string lcuenta2 = "";
            string lcuenta3 = "";
            string lcuenta4 = "";
            
            //checar el combobox de filtro para ver si es por saldos
            // si es por salods
            int lindice = -1;
            if (comboBox3.SelectedItem.ToString() == "Saldos No Cero")
            {
                foreach (DataRow row in saldos.Rows)
                {
                    lindice++;
                    if (lcuenta1 != row[1].ToString().Trim())
                    {
                        if (decimal.Parse(row[4].ToString().Trim()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[1].ToString().Trim();
                            sheet.Cells[lrenglon, 4].value = row[2].ToString().Trim();
                            sheet.Cells[lrenglon, 5].value = row[3].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][2].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][2].ToString();
                            sheet.Cells[lrenglon, 8].value = row[4].ToString().Trim();
                            lcuenta1 = row[1].ToString().Trim();
                            lrenglon++;
                        }
                    }

                    if (lcuenta2 != row[6].ToString().Trim())
                    {
                        if (decimal.Parse(row[9].ToString().Trim()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[6].ToString().Trim(); // cuenta
                            sheet.Cells[lrenglon, 4].value = row[7].ToString().Trim(); // nombre cuenta
                            sheet.Cells[lrenglon, 5].value = row[8].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][5].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][5].ToString();
                            sheet.Cells[lrenglon, 8].value = row[9].ToString().Trim();
                            lcuenta2 = row[6].ToString().Trim();

                            lrenglon++;
                        }
                    }
                    if (lcuenta3 != row[11].ToString().Trim())
                    {
                        if (decimal.Parse(row[14].ToString().Trim()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[11].ToString().Trim(); // cuenta
                            sheet.Cells[lrenglon, 4].value = row[12].ToString().Trim(); // nombre cuenta
                            sheet.Cells[lrenglon, 5].value = row[13].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][8].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][8].ToString();
                            sheet.Cells[lrenglon, 8].value = row[14].ToString().Trim();
                            lcuenta3 = row[11].ToString().Trim();
                            lrenglon++;
                        }
                    }

                    if (int.Parse(row[15].ToString()) != 0)
                    {
                        if (lcuenta4 != row[16].ToString().Trim())
                        {
                            if (decimal.Parse(row[19].ToString().Trim()) != 0)
                            {
                                sheet.Cells[lrenglon, 3].value = "'" + row[16].ToString().Trim(); // cuenta
                                sheet.Cells[lrenglon, 4].value = row[17].ToString().Trim(); // nombre cuenta
                                sheet.Cells[lrenglon, 5].value = row[18].ToString().Trim();
                                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][11].ToString();
                                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][11].ToString();
                                sheet.Cells[lrenglon, 8].value = row[19].ToString().Trim();
                                lcuenta4 = row[16].ToString().Trim();
                                lrenglon++;
                            }
                        }
                        if (int.Parse(row[20].ToString()) != 0)
                            if (decimal.Parse(row[24].ToString().Trim()) != 0)
                            {
                                sheet.Cells[lrenglon, 3].value = "'" + row[21].ToString().Trim(); // cuenta
                                sheet.Cells[lrenglon, 4].value = row[22].ToString().Trim(); // nombre cuenta
                                sheet.Cells[lrenglon, 5].value = row[23].ToString().Trim();
                                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][14].ToString();
                                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][14].ToString();
                                sheet.Cells[lrenglon, 8].value = row[24].ToString().Trim();

                                lrenglon++;
                            }
                    }

                }
            }
            else
            {
                foreach (DataRow row in saldos.Rows)
                {
                    lindice++;
                    if (lcuenta1 != row[1].ToString().Trim())
                    {
                        if (decimal.Parse(cargos.Rows[lindice][2].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][2].ToString()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[1].ToString().Trim();
                            sheet.Cells[lrenglon, 4].value = row[2].ToString().Trim();
                            sheet.Cells[lrenglon, 5].value = row[3].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][2].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][2].ToString();
                            sheet.Cells[lrenglon, 8].value = row[4].ToString().Trim();
                            lcuenta1 = row[1].ToString().Trim();
                            lrenglon++;
                        }
                    }

                    if (lcuenta2 != row[6].ToString().Trim())
                    {
                        if (decimal.Parse(cargos.Rows[lindice][5].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][5].ToString()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[6].ToString().Trim(); // cuenta
                            sheet.Cells[lrenglon, 4].value = row[7].ToString().Trim(); // nombre cuenta
                            sheet.Cells[lrenglon, 5].value = row[8].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][5].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][5].ToString();
                            sheet.Cells[lrenglon, 8].value = row[9].ToString().Trim();
                            lcuenta2 = row[6].ToString().Trim();

                            lrenglon++;
                        }
                    }
                    if (lcuenta3 != row[11].ToString().Trim())
                    {
                        if (decimal.Parse(cargos.Rows[lindice][8].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][8].ToString()) != 0)
                        {
                            sheet.Cells[lrenglon, 3].value = "'" + row[11].ToString().Trim(); // cuenta
                            sheet.Cells[lrenglon, 4].value = row[12].ToString().Trim(); // nombre cuenta
                            sheet.Cells[lrenglon, 5].value = row[13].ToString().Trim();
                            sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][8].ToString();
                            sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][8].ToString();
                            sheet.Cells[lrenglon, 8].value = row[14].ToString().Trim();
                            lcuenta3 = row[11].ToString().Trim();
                            lrenglon++;
                        }
                    }

                    if (int.Parse(row[15].ToString()) != 0)
                    {
                        if (lcuenta4 != row[16].ToString().Trim())
                        {
                            if (decimal.Parse(cargos.Rows[lindice][11].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][11].ToString()) != 0)
                            {
                                sheet.Cells[lrenglon, 3].value = "'" + row[16].ToString().Trim(); // cuenta
                                sheet.Cells[lrenglon, 4].value = row[17].ToString().Trim(); // nombre cuenta
                                sheet.Cells[lrenglon, 5].value = row[18].ToString().Trim();
                                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][11].ToString();
                                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][11].ToString();
                                sheet.Cells[lrenglon, 8].value = row[19].ToString().Trim();
                                lcuenta4 = row[16].ToString().Trim();
                                lrenglon++;
                            }
                        }
                        if (int.Parse(row[20].ToString()) != 0)
                            if (decimal.Parse(cargos.Rows[lindice][14].ToString()) != 0 || decimal.Parse(abonos.Rows[lindice][14].ToString()) != 0)
                            {
                                sheet.Cells[lrenglon, 3].value = "'" + row[21].ToString().Trim(); // cuenta
                                sheet.Cells[lrenglon, 4].value = row[22].ToString().Trim(); // nombre cuenta
                                sheet.Cells[lrenglon, 5].value = row[23].ToString().Trim();
                                sheet.Cells[lrenglon, 6].value = cargos.Rows[lindice][14].ToString();
                                sheet.Cells[lrenglon, 7].value = abonos.Rows[lindice][14].ToString();
                                sheet.Cells[lrenglon, 8].value = row[24].ToString().Trim();

                                lrenglon++;
                            }
                    }

                }

            }
            sheet.Cells.EntireColumn.AutoFit();


        }

        private void mTraerInfo(ref DataSet info)
        {
            int aindicesaldomenor=0;
            int aindicesaldomayor=0;
            aindicesaldomayor = comboBox1.SelectedIndex + 1;

            int lejercicio = comboBox2.SelectedIndex;
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
            if (lindicesaldomenor =="0")
                lindicesaldomenor = "";
            string lindicesaldomayor = aindicesaldomayor.ToString();


            

            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);
            SqlCommand mySqlCommand = new SqlCommand();
            string saldos = "SELECT c1.id, c1.codigo, c1.Nombre, isnull(sc1." + lcampoinicial + lindicesaldomenor + ",0), isnull(sc1.Importes" + lindicesaldomayor + ",0), " +
                  " c2.id, c2.codigo, c2.Nombre, isnull(sc2." + lcampoinicial + lindicesaldomenor + ",0), isnull(sc2.Importes" + lindicesaldomayor + ",0), " +
                  " c3.id, c3.codigo, c3.Nombre, isnull(sc3." + lcampoinicial + lindicesaldomenor + ",0), isnull(sc3.Importes" + lindicesaldomayor + ",0), " +
                  " isnull(c4.id,0), c4.codigo, c4.Nombre, isnull(sc4." + lcampoinicial + lindicesaldomenor + ",0), isnull(sc4.Importes" + lindicesaldomayor + ",0), " +
                  " isnull(c5.id,0), c5.codigo, c5.Nombre, isnull(sc5." + lcampoinicial + lindicesaldomenor + ",0), isnull(sc5.Importes" + lindicesaldomayor + ",0) " +
                  " FROM CUENTAS c1  " +
                  " join Asociaciones a1 on c1.Id = a1.IdCtaSup " +
                  " join saldoscuentas sc1 on c1.id = sc1.IdCuenta and sc1.tipo = 1 " +
                  " join Ejercicios e1 on e1.id = sc1.Ejercicio and e1.Ejercicio = 2014  " +
                  " join cuentas c2 on c2.id = a1.IdSubCtade  " +
                  " join saldoscuentas sc2 on c2.id = sc2.IdCuenta and sc2.tipo = 1 " +
                  " join Ejercicios e2 on e2.id = sc2.Ejercicio and e2.Ejercicio = 2014  " +
                  " join Asociaciones a2 on c2.Id = a2.IdCtaSup " +
                  " left join cuentas c3 on c3.id = a2.IdSubCtade  " +
                  " left join saldoscuentas sc3 on c3.id = sc3.IdCuenta and sc3.tipo = 1  " +
                  " join Ejercicios e3 on e3.id = sc3.Ejercicio and e3.Ejercicio = 2014  " +
                  " left join Asociaciones a3 on c3.Id = a3.IdCtaSup " +
                  " left join cuentas c4 on c4.id = a3.IdSubCtade  " +
                  " left join saldoscuentas sc4 on c4.id = sc4.IdCuenta and sc4.tipo = 1 and sc4.ejercicio = 5  " +
                  " left join Asociaciones a4 on c4.Id = a4.IdCtaSup " +
                  " left join cuentas c5 on c5.id = a4.IdSubCtade  " +
                  " left join saldoscuentas sc5 on c5.id = sc5.IdCuenta and sc5.tipo = 1 and sc5.ejercicio = 5 " +
                  " WHERE c1.CtaMayor =4 " +
                  " order by a1.Id, a2.id, a3.id, a4.id;  ";

            saldos += " SELECT c1.id, c1.Nombre, isnull(sc1.Importes" + lindicesaldomayor + ",0), " +
                    " c2.id, c2.Nombre, isnull(sc2.Importes" + lindicesaldomayor + ",0), " +
                    " c3.id, c3.Nombre, isnull(sc3.Importes" + lindicesaldomayor + ",0), " +
                    " c4.id, c4.Nombre, isnull(sc4.Importes" + lindicesaldomayor + ",0), " +
                    " c5.id, c5.Nombre, isnull(sc5.Importes" + lindicesaldomayor + ",0) " +
                    " FROM CUENTAS c1  "+
                    " join Asociaciones a1 on c1.Id = a1.IdCtaSup "+
                    " join saldoscuentas sc1 on c1.id = sc1.IdCuenta and sc1.tipo = 2 "+
                    " join Ejercicios e1 on e1.id = sc1.Ejercicio and e1.Ejercicio = 2014  "+
                    " join cuentas c2 on c2.id = a1.IdSubCtade  "+
                    " join saldoscuentas sc2 on c2.id = sc2.IdCuenta and sc2.tipo = 2 "+
                    " join Ejercicios e2 on e2.id = sc2.Ejercicio and e2.Ejercicio = 2014  "+
                    " join Asociaciones a2 on c2.Id = a2.IdCtaSup "+
                    " left join cuentas c3 on c3.id = a2.IdSubCtade  "+
                    " left join saldoscuentas sc3 on c3.id = sc3.IdCuenta and sc3.tipo = 2 "+
                    " join Ejercicios e3 on e3.id = sc3.Ejercicio and e3.Ejercicio = 2014  "+
                    " left join Asociaciones a3 on c3.Id = a3.IdCtaSup "+
                    " left join cuentas c4 on c4.id = a3.IdSubCtade  " +
                    " left join saldoscuentas sc4 on c4.id = sc4.IdCuenta and sc4.tipo = 2 and sc4.ejercicio = 5  "+
                    " left join Asociaciones a4 on c4.Id = a4.IdCtaSup "+
                    " left join cuentas c5 on c5.id = a4.IdSubCtade  "+
                    " left join saldoscuentas sc5 on c5.id = sc5.IdCuenta and sc5.tipo = 2 and sc5.ejercicio = 5  "+
                    " WHERE c1.CtaMayor =4 "+
                    " order by a1.Id, a2.id, a3.id, a4.id; ";

            saldos += " SELECT c1.id, c1.Nombre, isnull(sc1.Importes" + lindicesaldomayor + ",0), " +
                    " c2.id, c2.Nombre, isnull(sc2.Importes" + lindicesaldomayor + ",0), " +
                    " c3.id, c3.Nombre, isnull(sc3.Importes" + lindicesaldomayor + ",0), " +
                    " c4.id, c4.Nombre, isnull(sc4.Importes" + lindicesaldomayor + ",0), " +
                    " c5.id, c5.Nombre, isnull(sc5.Importes" + lindicesaldomayor + ",0) " +
                    " FROM CUENTAS c1  " +
                    " join Asociaciones a1 on c1.Id = a1.IdCtaSup " +
                    " join saldoscuentas sc1 on c1.id = sc1.IdCuenta and sc1.tipo = 3 " +
                    " join Ejercicios e1 on e1.id = sc1.Ejercicio and e1.Ejercicio = 2014  " +
                    " join cuentas c2 on c2.id = a1.IdSubCtade  " +
                    " join saldoscuentas sc2 on c2.id = sc2.IdCuenta and sc2.tipo = 3 " +
                    " join Ejercicios e2 on e2.id = sc2.Ejercicio and e2.Ejercicio = 2014  " +
                    " join Asociaciones a2 on c2.Id = a2.IdCtaSup " +
                    " left join cuentas c3 on c3.id = a2.IdSubCtade  " +
                    " left join saldoscuentas sc3 on c3.id = sc3.IdCuenta and sc3.tipo = 3 " +
                    " join Ejercicios e3 on e3.id = sc3.Ejercicio and e3.Ejercicio = 2014  " +
                    " left join Asociaciones a3 on c3.Id = a3.IdCtaSup " +
                    " left join cuentas c4 on c4.id = a3.IdSubCtade  " +
                    " left join saldoscuentas sc4 on c4.id = sc4.IdCuenta and sc4.tipo = 3 and sc4.ejercicio = 5  " +
                    " left join Asociaciones a4 on c4.Id = a4.IdCtaSup " +
                    " left join cuentas c5 on c5.id = a4.IdSubCtade  " +
                    " left join saldoscuentas sc5 on c5.id = sc5.IdCuenta and sc5.tipo = 3 and sc5.ejercicio = 5  " +
                    " WHERE c1.CtaMayor =4 " +
                    " order by a1.Id, a2.id, a3.id, a4.id ";

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
