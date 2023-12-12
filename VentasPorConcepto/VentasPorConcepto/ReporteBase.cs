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
using MyExcel = Microsoft.Office.Interop.Excel;


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
                    //MessageBox.Show("Conexion correcta, volver a ejecutar el exe");
                    //this.Close();
                }
            }
        }
        protected void mEncabezadoCelda(MyExcel.Worksheet sheet, string inicio, string fin, int lrenglon, int lagregarenglon, int tamano, string texto, Boolean Bold = true)
        {

            sheet.Cells[lrenglon, fin].value = texto;


            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Merge();

            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).HorizontalAlignment = MyExcel.XlHAlign.xlHAlignCenter;
            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).VerticalAlignment = MyExcel.XlHAlign.xlHAlignCenter;

            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Font.Size = tamano;

            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Font.Bold = Bold;

            sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).WrapText = true;
        }
    }
}
