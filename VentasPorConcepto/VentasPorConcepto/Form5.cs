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
    public partial class Form5 : Form
    {

        VentasSegmentoNegocio y = new VentasSegmentoNegocio();
        Antiguedad z = new Antiguedad();


        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = Properties.Settings.Default.database;
            txtBD.Text = "CompacwAdmin";
            txtBD.Enabled = false;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;

            txtServerC.Text = Properties.Settings.Default.server2;
            txtBD.Enabled = false;
            txtUsuarioC.Text = Properties.Settings.Default.user2;
            txtPasswordC.Text = Properties.Settings.Default.password2;

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

        private bool mValidaC()
        {
            string Cadenaconexion = "data source =" + txtServerC.Text + ";initial catalog =" + txtBDC.Text + ";user id = " + txtUsuarioC.Text + "; password = " + txtPasswordC.Text + ";";
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
            if (mValidaC() && mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.server2 = txtServerC.Text;
                Properties.Settings.Default.database2 = txtBDC.Text;
                Properties.Settings.Default.user2 = txtUsuarioC.Text;
                Properties.Settings.Default.password2 = txtPasswordC.Text;


                Properties.Settings.Default.Save();

                this.Close();
                this.DialogResult = DialogResult.OK;
                y.Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                //y.mllenarcomboempresas();

                z.Cadenaconexion = "data source =" + Properties.Settings.Default.server2 +
                ";initial catalog =" + Properties.Settings.Default.database2 + " ;user id = " + Properties.Settings.Default.user2 +
                "; password = " + Properties.Settings.Default.password2 + ";";
                y.Visible = true;
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        
    }

        private void button2_Click(object sender, EventArgs e)
        {
            if (mValidaC() && mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.server2 = txtServerC.Text;
                Properties.Settings.Default.database2 = txtBDC.Text;
                Properties.Settings.Default.user2 = txtUsuarioC.Text;
                Properties.Settings.Default.password2 = txtPasswordC.Text;


                Properties.Settings.Default.Save();

                this.Close();
                this.DialogResult = DialogResult.OK;
                y.Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                //y.mllenarcomboempresas();

                z.Cadenaconexion = "data source =" + Properties.Settings.Default.server2+
                ";initial catalog =" + Properties.Settings.Default.database2 + " ;user id = " + Properties.Settings.Default.user2 +
                "; password = " + Properties.Settings.Default.password2 + ";";
                y.Visible = true;
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }
    }
}
