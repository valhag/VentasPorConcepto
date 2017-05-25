using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace VentasPorConcepto
{

    /*public class RegConcepto1
    {
        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _sTipocfd;

        public string Tipocfd
        {
            get { return _sTipocfd; }
            set { _sTipocfd = value; }
        }
        private long _id;

        public long id
        {
            get { return _id; }
            set { _id = value; }
        }

    }*/
    public partial class InventarioCapas : Form
    {
        Class1 x = new Class1();

        List<RegConcepto> listavalores;

        public InventarioCapas()
        {
            InitializeComponent();
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

                lfiltroproducto = " and (m5.ccodigop01 >= '" + textBox1.Text + "' and  m5.ccodigop01 <= '" +  textBox2.Text + "') ";
                lfiltroproducto1 = " where (m5.ccodigop01 >= '" + textBox1.Text + "' and  m5.ccodigop01 <= '" +  textBox2.Text + "') ";
            }

            string lfiltroalmacen = "";
            if (textBox3.Text != "" && textBox4.Text != "")
                lfiltroalmacen = " and (m3.ccodigoa01 >= '" + textBox3.Text + "' and  m3.ccodigoa01 <= '" + textBox4.Text + "') ";

            string lfiltroclasif = "";
            if (listBox2.Items.Count > 0)
            { 
                foreach (Class1.RegConcepto element in listBox2.Items)
                {
                    if (lfiltroclasif == "")
                        lfiltroclasif += " and (";
                    else
                        lfiltroclasif += " or ";
                    lfiltroclasif += " m20a.cidvalor01 = " + element.id;
                }
                lfiltroclasif += ")";

            
            }


            /*if (textBox3.Text != "" && textBox4.Text != "")
                lfiltroproducto = " and (m5.ccodigop01 >= '" + textBox1.Text + "' and  m5.ccodigop01 <= '" + textBox2.Text + "') ";*/

            /*lfiltroproducto = "and m5.ccodigop01 = 'CDV001'";*/


            string lquery0 = "select m5.cidprodu01 as idprodue, m5.ccodigop01, m5.cnombrep01 as nombree, m5.cmetodoc01, m20a.cvalorcl01, m20b.cvalorcl01,m20c.cvalorcl01,m20d.cvalorcl01,m20e.cvalorcl01,m20f.cvalorcl01 " +
            " from mgw10005 m5 " +
            " join mgw10020 m20a on m20a.cidvalor01 = m5.cidvalor01 " + lfiltroclasif +
" join mgw10020 m20b on m20b.cidvalor01 = m5.cidvalor02 " + 
" join mgw10020 m20c on m20c.cidvalor01 = m5.cidvalor03 " +
" join mgw10020 m20d on m20d.cidvalor01 = m5.cidvalor04 " +
" join mgw10020 m20e on m20e.cidvalor01 = m5.cidvalor05 " +
" join mgw10020 m20f on m20f.cidvalor01 = m5.cidvalor06 " + lfiltroproducto1;
            lista.Add(lquery0);

            string lquery1 = "select m5.cidprodu01 as idprodue, m5.ccodigop01, m5.cnombrep01 as nombree, m5.cmetodoc01, sum(m10.cunidades) as unie, m20a.cvalorcl01, m20b.cvalorcl01,m20c.cvalorcl01,m20d.cvalorcl01,m20e.cvalorcl01,m20f.cvalorcl01 " +
            " from mgw10010 m10 join mgw10005 m5 " +
" on m10.cidprodu01 = m5.cidprodu01 " + lfiltroproducto +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" join mgw10020 m20a on m20a.cidvalor01 = m5.cidvalor01 " + lfiltroclasif +
" join mgw10020 m20b on m20b.cidvalor01 = m5.cidvalor02 " + 
" join mgw10020 m20c on m20c.cidvalor01 = m5.cidvalor03 " +
" join mgw10020 m20d on m20d.cidvalor01 = m5.cidvalor04 " +
" join mgw10020 m20e on m20e.cidvalor01 = m5.cidvalor05 " +
" join mgw10020 m20f on m20f.cidvalor01 = m5.cidvalor06 " + 
" where dtos(m10.cfecha) < '" + sfecha1 + "' and m10.cafectad02 =1 and m10.cafectae01 =1 " +
" group by m5.cidprodu01, m5.ccodigop01, m5.cnombrep01,m5.cmetodoc01, m20a.cvalorcl01, m20b.cvalorcl01,m20c.cvalorcl01,m20d.cvalorcl01,m20e.cvalorcl01,m20f.cvalorcl01 ";


            lquery1 = "select m5.cidprodu01 as idprodue, sum(m10.cunidades) as unie" +
            " from mgw10010 m10 join mgw10005 m5 " +
" on m10.cidprodu01 = m5.cidprodu01 " + lfiltroproducto +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where dtos(m10.cfecha) < '" + sfecha1 + "' and m10.cafectad02 =1 and m10.cafectae01 =1 " +
" group by m5.cidprodu01";

            // inventario inicial entradas
            lista.Add(lquery1);

            string lquery2 = "select m5.cidprodu01 as idprodus, m5.ccodigop01, m5.cnombrep01 as nombree, m5.cmetodoc01, sum(m10.cunidades) as unis from mgw10010 m10 join mgw10005 m5 " +
" on m10.cidprodu01 = m5.cidprodu01 " + lfiltroproducto +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where dtos(m10.cfecha) < '" + sfecha1 + "' and m10.cafectad02 =1 and m10.cafectae01 =2" +
" group by m5.cidprodu01, m5.ccodigop01, m5.cnombrep01,m5.cmetodoc01 ";

            // inventario inicial salidas
            lista.Add(lquery2);

            // capas inventario inicial entradas

            
            //lfiltroproducto = "";

            string lquery3 = "select m10.cidprodu01, m28.cidcapa, dtos(m28.cfecha) as cfecha, m28.cunidades, m25.ccosto, m3.cnombrea01 from mgw10010 m10" +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01" +
" join mgw10028 m28 on m10.cidmovim01 = m28.cidmovim01" +
" join mgw10025 m25 on m25.cidcapa = m28.cidcapa" +
" join mgw10003 m3 on m3.cidalmacen = m25.cidalmacen" + lfiltroalmacen +
" where m10.cafectad02 = 1 and m10.cafectae01 =1 and dtos(m10.cfecha) < '" + sfecha1 + "'" + lfiltroproducto +
" order by m10.cidprodu01 ";

lista.Add(lquery3);

// capas inventario inicial salidas
string lquery4 = "select m10.cidprodu01, m28.cidcapa, sum(m28.cunidades) from mgw10010 m10 " +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01 " +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" join mgw10028 m28 on m10.cidmovim01 = m28.cidmovim01 " +
" join mgw10025 m25 on m25.cidcapa = m28.cidcapa " +
" where m10.cafectad02 = 1 and m10.cafectae01 =2 and dtos(m10.cfecha) < '" + sfecha1 + "'"  + lfiltroproducto +  
" group by m10.cidprodu01, m28.cidcapa ";
lista.Add(lquery4);

// movimientos periodo entrada
string lquery5 = "select m5.cidprodu01 as idprodue, sum(m10.cunidades) as unie from mgw10010 m10 join mgw10005 m5 " +
" on m10.cidprodu01 = m5.cidprodu01  " +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where dtos(m10.cfecha) >= '" + sfecha1 + "' and dtos(m10.cfecha) <= '" + sfecha2 + "' and m10.cafectad02 =1 and m10.cafectae01 =1 " + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
" group by m5.cidprodu01 ";
lista.Add(lquery5);

// movimientos periodo salida

string lquery6 = "select m5.cidprodu01 as idprodus, sum(m10.cunidades) as unie from mgw10010 m10 join mgw10005 m5 " +
" on m10.cidprodu01 = m5.cidprodu01  " +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" where dtos(m10.cfecha) >= '" + sfecha1 + "' and dtos(m10.cfecha) <= '" + sfecha2 + "' and m10.cafectad02 =1 and m10.cafectae01 =2 " + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
" group by m5.cidprodu01 ";
lista.Add(lquery6);


// capas en periodo entrada
string lquery7 = "select m10.cidprodu01, m28.cidcapa,  m28.cunidades,m25.ccosto, m3.cnombrea01 from mgw10010 m10" + // dtos(m28.cfecha) as cfecha,
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01" +
" join mgw10028 m28 on m10.cidmovim01 = m28.cidmovim01" +
" join mgw10025 m25 on m25.cidcapa = m28.cidcapa" +
" join mgw10003 m3 on m3.cidalmacen = m25.cidalmacen" + lfiltroalmacen +
" where m10.cafectad02 = 1 and m10.cafectae01 =1 and dtos(m10.cfecha) >= '" + sfecha1 + "' and dtos(m10.cfecha) <= '" + sfecha2 + "'" + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
" order by m10.cidprodu01 ";
lista.Add(lquery7);

// capas en periodo salida
string lquery8 = "select m10.cidprodu01, m28.cidcapa, sum(m28.cunidades) from mgw10010 m10 " +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01 " +
" join mgw10003 m3 on m3.cidalmacen = m10.cidalmacen " + lfiltroalmacen +
" join mgw10028 m28 on m10.cidmovim01 = m28.cidmovim01 " +
" join mgw10025 m25 on m25.cidcapa = m28.cidcapa " +
" where m10.cafectad02 = 1 and m10.cafectae01 = 2 and dtos(m10.cfecha) >= '" + sfecha1 + "' and dtos(m10.cfecha) <= '" + sfecha2 + "'" + lfiltroproducto + // and m5.ccodigop01 = 'CAJ001'" + 
" group by m10.cidprodu01, m28.cidcapa ";
lista.Add(lquery8);




//            if (catalogo1.mRegresarCodigo() != "" && catalogo2.mRegresarCodigo() != "")
  //              lquery += " and m2.ccodigoc01 >= '" + catalogo1.mRegresarCodigo() + "' and m2.ccodigoc01 <= '" + catalogo2.mRegresarCodigo() + "'";



    //        lquery += " order by m8.cfecha, m8.cfolio";


            x.mTraerDataset(lista, comboBox1.SelectedValue.ToString());

            string lfechai = dateTimePicker1.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker1.Value.Year.ToString().PadLeft(4, '0');
            string lfechaf = dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + '/' + dateTimePicker2.Value.Year.ToString().PadLeft(4, '0');


            x.mReporteInventarioCapas(comboBox1.Text, lfechai, lfechaf);
        }

        public OleDbConnection mAbrirConexionOrigen(string mEmpresa)
        {
            OleDbConnection _conexion;
            _conexion = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Server=localhost;Database=adMazorca;User Id=sa;Password=ady123";
                _conexion.Open();
            }
            return _conexion;

        }

        private void InventarioCapas_Load(object sender, EventArgs e)
        {
            this.Text = "Inventario Capas";
            string mensaje;
            this.comboBox1.Items.Clear();
            this.comboBox1.DataSource = x.mCargarEmpresas(out mensaje);
            comboBox1.DisplayMember = "Nombre";
            comboBox1.ValueMember = "Ruta";
            try
            {
                this.comboBox1.SelectedIndex = 1;
                this.comboBox1.SelectedIndex = 0;
            }
            catch (Exception ee)
            {
                this.comboBox1.SelectedIndex = 0;
            }

            //List<RegConcepto1> listavalores1 = new List<RegConcepto1>();

            //List<object> collection = new List<RegConcepto>((IEnumerable)RegConcepto);
            
            x.mCargarClasificaciones(comboBox1.SelectedValue.ToString(), 1);


            //listof< RegConcepto


            listBox1.DataSource = x._RegClasificaciones;
            //listBox1.DataSource = listavalores;
            listBox1.DisplayMember = "Nombre";
            listBox1.ValueMember = "Id";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            x.mCargarClasificaciones(comboBox1.SelectedValue.ToString(), 1);
            listBox1.DataSource = null;
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox1.DataSource = x._RegClasificaciones;
            //listBox1.DataSource = listavalores;
            
            listBox1.DisplayMember = "Nombre";
            listBox1.ValueMember = "Id";
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
        private void button2_Click(object sender, EventArgs e)
        {
            Asignar(listBox1, listBox2);
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            deAsignar();
        }
    }
}
