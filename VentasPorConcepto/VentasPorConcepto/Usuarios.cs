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
    public partial class Usuarios : Form
    {
        string Cadenaconexion = "";
        Class1 x = new Class1();
        public Usuarios()
        {
            InitializeComponent();
        }

        private void Usuarios_Load(object sender, EventArgs e)
        {
            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = "RepositorioAdminpaq";
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = "RepositorioAdminpaq";
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();

                //this.Close();
               Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                //y.mllenarcomboempresas();
                //y.Visible = true;
               MessageBox.Show("Valores de conexion correctos");
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

        private void button2_Click(object sender, EventArgs e)
        {
            string lquery;

            // string archivo = @"C:\fromgithub\archivotest.xlsx";

            string archivo = @textBox1.Text;

            lquery = "SELECT C1.IDUSUARIO, CLAVE AS NOMBRE " +
",NOMBRE AS NOMBRELARGO, C3.DESCRIPCION AS GRUPO, C2H.PROCESO " +
"FROM CAC10000 C1 " +
"JOIN CAC30000 C3 ON C1.NIVEL = C3.IDPERFIL " +
"join CAC20000 C2H on c1.NIVEL = C2H.NIVEL " +
"AND CHARINDEX('E',LTRIM(C2H.PROCESO),0) > 0 " +
"AND C2H.ESTADO = 1  " +
"ORDER BY C1.IDUSUARIO;" +
"select C1.IDUSUARIO, c4.DESCRIPCION as proceso, c2.ESTADO, c5.DESCRIPCION as grupo " +
"from CAC10000 c1 " +
"join CAC30000 c3 on c1.NIVEL = c3.IDPERFIL " +
"join CAC20000 c2 on c3.IDPERFIL = c2.NIVEL " +
"join CAC40000 c4 on c4.IDPROCESO = c2.PROCESO " +
"join CAC50000 c5 on c5.GRUPO = c4.GRUPO " +
"where c1.IDSISTEMA =5  ";

            if (textBox1.Text != "")
            {
                lquery += " and c1.clave = '" + textBox1.Text + "'";
            }

lquery += "order by C1.IDUSUARIO, C4.GRUPO,C4.CIDAUTOINCSQL " ;






            x.mTraerInformacionComercial2(lquery, txtBD.Text);



         x.mReporteUsuarios();
            //   x.mTestFotos();
        }
    }
}
