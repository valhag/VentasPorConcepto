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
    public partial class ReporteBase : Form
    {
        public ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        //public Form4 y = new Form4();
        public Class1 x = new Class1();

        public ReporteBase()
        {
            //y.Visible = false;
            InitializeComponent();
        }

        private void ReporteBase_Load(object sender, EventArgs e)
        {
            //y.Visible = false;
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
                Form4 y = new Form4();
                y.Visible = false;
                DialogResult lresp = y.ShowDialog(this);
                if (lresp == DialogResult.OK)
                {
                    Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                    ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                    "; password = " + Properties.Settings.Default.password + ";";
                    empresasComercial1.Populate(Cadenaconexion);
                }
            }
        }
    }
}
