﻿using System;
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
    public partial class Form4 : Form
    {
        // Form1 y = new Form1();
        //FacturaPedidoComercial y = new FacturaPedidoComercial();
        //CapasComercial y = new CapasComercial();

        //AmecaConceptos z = new AmecaConceptos();

        //Form1 y = new Form1();
        //Gomar z = new Gomar();

        //public void asignaform1(Form1 ay)
        //{
        //    y = ay;
        //}

        //public void asignaTLS(TLS ay)
        //{
        //    y = new TLS();
        //    y = ay;
        //}
        //ReporteBase miForm = new ReporteBase();
        
        public Form4()
        {
            this.Visible = false;
            InitializeComponent();
        }

        /*public Form4(Form x)
        {
            miForm = (ReporteBase)x;
            InitializeComponent();
        }*/

        private bool mValida()
        {
            string Cadenaconexion = "data source =" + txtServer.Text + ";initial catalog =" + txtBD.Text + ";user id = " + txtUser.Text + "; password = " + txtPass.Text  + ";";
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
            if (mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();
                this.DialogResult = DialogResult.OK;
                this.Close();
                
                
                /*miForm.Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";*/
                //y.mllenarcomboempresas();
                //this.Parent.Visible = true;
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }

        

        private void Form4_Load(object sender, EventArgs e)
        {
            txtServer.Text = Properties.Settings.Default.server  ;
            txtBD.Text = Properties.Settings.Default.database;
            txtBD.Text = "CompacwAdmin";
            txtBD.Enabled = false;
            txtUser.Text = Properties.Settings.Default.user ;
            txtPass.Text = Properties.Settings.Default.password ;
        }
    }
}
