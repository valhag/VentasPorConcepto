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
    public partial class CapasComercial : Form
    {

        ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
        Class1 x = new Class1();

        public CapasComercial()
        {
            InitializeComponent();
        }

        private void CapasComercial_Load(object sender, EventArgs e)
        {
            this.Text = " Reporte Capas Comercial " + " " + this.ProductVersion;
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());


            string server = Properties.Settings.Default.server;
            //MessageBox.Show("server " + server);
            if (Properties.Settings.Default.server != "")
            {

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + "CompacWAdmin" + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                //Archivo = Properties.Settings.Default.archivo;
            }
            if (Cadenaconexion != "")
            {
                empresasComercial1.Populate(Cadenaconexion);
                StringBuilder lquery = new StringBuilder();
                lquery.AppendLine("select CIDVALORCLASIFICACION, CCODIGOVALORCLASIFICACION, CVALORCLASIFICACION from admClasificacionesValores where CIDCLASIFICACION = 25");
                x.mTraerInformacionClasificacionesComercial(lquery,empresasComercial1.aliasbdd);
                //x.mCargarClasificacionesComercial(comboBox1.SelectedValue.ToString(), 1);


                //listof< RegConcepto


                listBox1.DataSource = x._RegClasificaciones;
                //listBox1.DataSource = listavalores;
                listBox1.DisplayMember = "Nombre";
                listBox1.ValueMember = "id";
            }
            else
            {
                Form4 x = new Form4();
                x.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Asignar(listBox1, listBox2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            deAsignar();
        }

        private void Asignar(ListBox listade, ListBox listaa)
        {
            List<Class1.RegConcepto> listaseleccionados = new List<Class1.RegConcepto>();
            foreach (Class1.RegConcepto element in listade.SelectedItems)
            {
                listaseleccionados.Add(element);
            }
            List<int> selectedItemIndexes = new List<int>();
            foreach (object o in listade.SelectedItems)
            {
                selectedItemIndexes.Add(listBox1.Items.IndexOf(o));
            }

            listade.DataSource = null;
            foreach (int i in selectedItemIndexes)
            {
                x._RegClasificaciones.RemoveAt(i);
            }
            listade.DataSource = x._RegClasificaciones;
            listade.DisplayMember = "Nombre";
            listade.ValueMember = "Id";

            foreach (Class1.RegConcepto element1 in listaseleccionados)
            {
                listaa.Items.Add(element1);
            }
            listaa.DisplayMember = "Nombre";
            listaa.ValueMember = "Id";
        }

        private void deAsignar()
        {
            List<Class1.RegConcepto> listaseleccionados = new List<Class1.RegConcepto>();
            foreach (Class1.RegConcepto element in listBox2.SelectedItems)
            {
                listaseleccionados.Add(element);
            }
            List<int> selectedItemIndexes = new List<int>();
            foreach (object o in listBox2.SelectedItems)
            {
                selectedItemIndexes.Add(listBox1.Items.IndexOf(o));
            }

            listBox1.DataSource = null;
            foreach (Class1.RegConcepto element1 in listaseleccionados)
            {
                x._RegClasificaciones.Add(element1);
            }
            listBox1.DataSource = x._RegClasificaciones;
            listBox1.DisplayMember = "Nombre";
            listBox1.ValueMember = "Id";

            foreach (Class1.RegConcepto element1 in listaseleccionados)
            {
                listBox2.Items.Remove(element1);
            }
            listBox2.DisplayMember = "Nombre";
            listBox2.ValueMember = "Id";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            string lquery;

            List<string> lista = new List<string>();

            string lfiltroproducto = "";
            string lfiltroproducto1 = "";
            if (textBox1.Text != "" && textBox2.Text != "")
            {

                lfiltroproducto = " and (m5.ccodigoproducto >= '" + textBox1.Text + "' and  m5.ccodigoproducto <= '" + textBox2.Text + "') ";
                lfiltroproducto1 = " where (m5.ccodigoproducto >= '" + textBox1.Text + "' and  m5.ccodigoproducto <= '" + textBox2.Text + "') ";
            }

            string lfiltroalmacen = "";
            if (textBox3.Text != "" && textBox4.Text != "")
                lfiltroalmacen = " and (m3.ccodigoalmacen >= '" + textBox3.Text + "' and  m3.ccodigoalmacen <= '" + textBox4.Text + "') ";

            string lfiltroclasif = "";
            if (listBox2.Items.Count > 0)
            {
                foreach (Class1.RegConcepto element in listBox2.Items)
                {
                    if (lfiltroclasif == "")
                        lfiltroclasif += " and (";
                    else
                        lfiltroclasif += " or ";
                    lfiltroclasif += " m20a.cidvalorclasificacion = " + element.id;
                }
                lfiltroclasif += ")";


            }


            /*if (textBox3.Text != "" && textBox4.Text != "")
                lfiltroproducto = " and (m5.ccodigop01 >= '" + textBox1.Text + "' and  m5.ccodigop01 <= '" + textBox2.Text + "') ";*/

            /*lfiltroproducto = "and m5.ccodigop01 = 'CDV001'";*/


            string lquery0 = "select m5.cidproducto as idprodue, m5.ccodigoproducto, m5.cnombreproducto as nombree, m5.cmetodocosteo, m20a.cvalorclasificacion, m20b.cvalorclasificacion,m20c.cvalorclasificacion,m20d.cvalorclasificacion,m20e.cvalorclasificacion,m20f.cvalorclasificacion " +
            " from admproductos m5 " +
            " join admClasificacionesValores m20a on m20a.cidvalorclasificacion = m5.cidvalorclasificacion1 " + lfiltroclasif +
" join admClasificacionesValores m20b on m20b.cidvalorclasificacion = m5.cidvalorclasificacion2 " +
" join admClasificacionesValores m20c on m20c.cidvalorclasificacion = m5.cidvalorclasificacion3 " +
" join admClasificacionesValores m20d on m20d.cidvalorclasificacion = m5.cidvalorclasificacion4 " +
" join admClasificacionesValores m20e on m20e.cidvalorclasificacion = m5.cidvalorclasificacion5 " +
" join admClasificacionesValores m20f on m20f.cidvalorclasificacion = m5.cidvalorclasificacion6 " + lfiltroproducto1;
            lista.Add(lquery0);

            string lquery1 = "select m5.cidproducto as idprodue, m5.ccodigoproducto, m5.cnombreproducto as nombree, m5.cmetodocosteo, sum(m10.cunidades) as unie, m20a.cvalorclasificacion, m20b.cvalorclasificacion,m20c.cvalorclasificacion,m20d.cvalorclasificacion,m20e.cvalorclasificacion,m20f.cvalorclasificacion " +
            " from admMovimientos m10 join admProductos m5 " +
" on m10.cidproducto = m5.cidproducto " + lfiltroproducto +
" join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" join admClasificacionesValores m20a on m20a.cidvalorclasificacion = m5.cidvalorclasificacion1 " + lfiltroclasif +
" join admClasificacionesValores m20b on m20b.cidvalorclasificacion = m5.cidvalorclasificacion2 " +
" join admClasificacionesValores m20c on m20c.cidvalorclasificacion = m5.cidvalorclasificacion3 " +
" join admClasificacionesValores m20d on m20d.cidvalorclasificacion = m5.cidvalorclasificacion4 " +
" join admClasificacionesValores m20e on m20e.cidvalorclasificacion = m5.cidvalorclasificacion5 " +
" join admClasificacionesValores m20f on m20f.cidvalorclasificacion = m5.cidvalorclasificacion6 " +
" where ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') < '" + sfecha1 + "' and m10.cafectadoinventario =1 and m10.cafectaexistencia =1 " +
" group by m5.cidproducto, m5.ccodigoproducto, m5.cnombreproducto,m5.cmetodocosteo, m20a.cvalorclasificacion, m20b.cvalorclasificacion,m20c.cvalorclasificacion,m20d.cvalorclasificacion,m20e.cvalorclasificacion,m20f.cvalorclasificacion ";


            lquery1 = "select m5.cidproducto as idprodue, sum(m10.cunidades) as unie" +
            " from admMovimientos m10 join admProductos m5 " +
" on m10.cidproducto = m5.cidproducto " + lfiltroproducto +
" join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  < '" + sfecha1 + "' and m10.cafectadoinventario =1 and m10.cafectaexistencia =1 " +
" group by m5.cidproducto";

            // inventario inicial entradas
            lista.Add(lquery1);

            string lquery2 = "select m5.cidproducto as idprodus, m5.ccodigoproducto, m5.cnombreproducto as nombree, m5.cmetodocosteo, sum(m10.cunidades) as unis from admMovimientos m10 join admProductos m5 " +
" on m10.cidproducto = m5.cidproducto " + lfiltroproducto +
" join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  < '" + sfecha1 + "' and m10.cafectadoinventario =1 and m10.cafectaexistencia =2" +
" group by m5.cidproducto, m5.ccodigoproducto, m5.cnombreproducto,m5.cmetodocosteo ";

            // inventario inicial salidas
            lista.Add(lquery2);

            // capas inventario inicial entradas


            //lfiltroproducto = "";

            string lquery3 = "select m10.cidproducto, m28.cidcapa, ltrim(str(year(m28.cfecha))) + REPLACE( ltrim(str(month(m28.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m28.cfecha),2)), SPACE(1), '0') as cfecha, m28.cunidades, m25.ccosto, m3.cnombrealmacen from admMovimientos m10" +
" join admProductos m5 on m5.cidproducto = m10.cidproducto" +
" join admMovimientosCapas m28 on m10.cidmovimiento = m28.cidmovimiento" +
" join admCapasProducto m25 on m25.cidcapa = m28.cidcapa" +
" join admAlmacenes m3 on m3.cidalmacen = m25.cidalmacen" + lfiltroalmacen +
" where m10.cafectadoinventario = 1 and m10.cafectaexistencia =1 and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  < '" + sfecha1 + "'" + lfiltroproducto +
" order by m10.cidproducto ";

            lista.Add(lquery3);

            // capas inventario inicial salidas
            string lquery4 = "select m10.cidproducto, m28.cidcapa, sum(m28.cunidades) from admMovimientos m10 " +
            " join admproductos m5 on m5.cidproducto = m10.cidproducto " +
            " join admalmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
            " join admmovimientoscapas m28 on m10.cidmovimiento = m28.cidmovimiento " +
            " join admcapasproducto m25 on m25.cidcapa = m28.cidcapa " +
            " where m10.cafectadoinventario = 1 and m10.cafectaexistencia =2 and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  < '" + sfecha1 + "'" + lfiltroproducto +
            " group by m10.cidproducto, m28.cidcapa ";
            lista.Add(lquery4);

            // movimientos periodo entrada
            string lquery5 = "select m5.cidproducto as idprodue, sum(m10.cunidades) as unie from admmovimientos m10 join admproductos m5 " +
            " on m10.cidproducto = m5.cidproducto  " +
            " join admalmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
            " where ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  >= '" + sfecha1 + "' and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') <= '" + sfecha2 + "' and m10.cafectadoinventario =1 and m10.cafectaexistencia =1 " + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
            " group by m5.cidproducto ";
            lista.Add(lquery5);

            // movimientos periodo salida

            string lquery6 = "select m5.cidproducto as idprodus, sum(m10.cunidades) as unie from admMovimientos m10 join admProductos m5 " +
            " on m10.cidproducto = m5.cidproducto  " +
            " join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
            " where ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  >= '" + sfecha1 + "' and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0')  <= '" + sfecha2 + "' and m10.cafectadoinventario =1 and m10.cafectaexistencia =2 " + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
            " group by m5.cidproducto ";
            lista.Add(lquery6);


            // capas en periodo entrada
            string lquery7 = "select m10.cidproducto, m28.cidcapa,  m28.cunidades,m25.ccosto, m3.cnombrealmacen from admmovimientos m10" + // dtos(m28.cfecha) as cfecha,
            " join admProductos m5 on m5.cidproducto = m10.cidproducto" +
            " join admMovimientosCapas m28 on m10.cidmovimiento = m28.cidmovimiento" +
            " join admcapasproducto m25 on m25.cidcapa = m28.cidcapa" +
            " join admAlmacenes m3 on m3.cidalmacen = m25.cidalmacen" + lfiltroalmacen +
            " where m10.cafectadoinventario = 1 and m10.cafectaexistencia =2 " +
            " and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') >= '" + sfecha1 + "' and ltrim(str(year(m10.cfecha))) + REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') <= '" + sfecha2 + "'" + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
            " order by m10.cidproducto ";
            lista.Add(lquery7);

            // capas en periodo salida
            string lquery8 = "select m10.cidproducto, m28.cidcapa, sum(m28.cunidades) from admMovimientos m10 " +
            " join admProductos m5 on m5.cidproducto = m10.cidproducto " +
            " join admAlmacenes m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
            " join admMovimientosCapas m28 on m10.cidmovimiento = m28.cidmovimiento " +
            " join admcapasproducto m25 on m25.cidcapa = m28.cidcapa " +
            " where m10.cafectadoinventario = 1 and m10.cafectaexistencia = 2 and REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') >= '" + sfecha1 + "' and REPLACE( ltrim(str(month(m10.cfecha),2)), SPACE(1), '0')  + REPLACE( ltrim(str(day(m10.cfecha),2)), SPACE(1), '0') <= '" + sfecha2 + "'" + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
            " group by m10.cidproducto, m28.cidcapa ";
            lista.Add(lquery8);



            x.mTraerDatasetComercial(lista, empresasComercial1.aliasbdd);

            string lfechai = dateTimePicker1.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Year.ToString().PadLeft(4, '0');
            string lfechaf = dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Year.ToString().PadLeft(4, '0');


            x.mReporteInventarioCapas(empresasComercial1.aliasbdd, lfechai, lfechaf);
        }
    }
}
