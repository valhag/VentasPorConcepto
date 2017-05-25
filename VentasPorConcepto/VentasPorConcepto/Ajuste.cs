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
    
    public partial class Ajuste : Form
    {

        private class Existencia
        {
            public decimal idproducto;
            public decimal unidades;
            //public int tipo;
        }

        public OleDbConnection _conexion;
        public DataSet Datos = null;

        public Ajuste()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            mCambiarCostos(textBox1.Text);
        }


        public void mCambiarCostos(string mEmpresa)
        {
            
            List<string> lqueries = new List<string>();
            lqueries.Clear();
            string lquery = "select cidprodu01, iif(m10.cafectae01 = 1,sum(cunidades), sum(cunidades)*-1), cafectae01 " +
            " from mgw10010 m10 where m10.cafectae01 < 3 and m10.cafectad01 = 1 group by cidprodu01, cafectae01 order by cidprodu01, cafectae01 ";

            lqueries.Add(lquery);

            lquery = "select m10.cidmovim01, m10.cidprodu01, m10.cunidades, m10.ciddocum01, m5.ccodigop01, m3.ccodigoa01, m10.cnumerom01 " +
            " from mgw10010 m10 join mgw10005 m5 on m10.cidprodu01 = m5.cidprodu01 " +
            " join mgw10003 m3 on m10.cidalmacen = m3.cidalmacen " +
            " where m10.cafectae01 = 1 and m10.cafectad01 = 1 order by m10.cidprodu01, m10.cfecha desc, m10.cafectae01 ";

            lqueries.Add(lquery);
            mTraerDataset(lqueries, mEmpresa);
            DataTable existencia = new DataTable();
            DataTable existencia1 = new DataTable();
            DataTable movtos = new DataTable();

            DataTable ExistenciaFinal = new DataTable("NameGroups");
            ExistenciaFinal.Columns.Add("Id", typeof(decimal));
            ExistenciaFinal.Columns.Add("Suma", typeof(decimal));

            existencia = Datos.Tables[0];
            movtos = Datos.Tables[1];


            var q = from row in existencia.AsEnumerable()
                    group row by row.Field<decimal>(0) into grp
                    select new
                    {
                        Id = grp.Key,
                        Sum = grp.Sum(r => r.Field<decimal>(1))
                    };

            foreach (var item in q)
            {
                ExistenciaFinal.Rows.Add(item.Id, item.Sum);
            }

            var combinado = from m in movtos.AsEnumerable()
                                    join e in ExistenciaFinal.AsEnumerable() on (string)m["cidprodu01"].ToString() equals (string)e["Id"].ToString()  //into tempp from e1
                            select new
                            {
                                IdMovim = m.Field<decimal>(0),
                                IdProdu = m.Field<decimal>(1),
                                UnidadesMovto = decimal.Parse(m.Field<double>(2).ToString()),
                                Existencia = e.Field<decimal>(1),
                                IdDocumento = decimal.Parse(m.Field<decimal>(3).ToString()),
                                CodigoProducto = m.Field<string>(4).ToString(),
                                CodigoAlmacen = m.Field<string>(5).ToString(),
                                NumeroRenglon = decimal.Parse(m.Field<double>(6).ToString())
                            };
            

            decimal totalexistencia = existencia.AsEnumerable().Sum(o => o.Field<decimal>(1));


            // 3114328 2399635     714693

            decimal costounitario = 2000000 / totalexistencia;
            decimal idproducto = 0;
            decimal lasignar = 0;
            int yadividido = 0;

            OleDbCommand com = new OleDbCommand();
            com.Connection = _conexion;
            foreach (var movto in combinado)
            {
                if (movto.IdProdu != idproducto)
                {
                    lasignar = movto.Existencia;
                }
                idproducto = movto.IdProdu;
                if (lasignar >0)
                {
                    if (lasignar > movto.UnidadesMovto )
                    {
                        // no romper solo cambiar el costo de la entrada 
                        com.CommandText = "update mgw10010 set ccostoca01 = ccostca01 + " + costounitario + " where cidmovim01 = " + movto.IdMovim;
                        int lafectados = com.ExecuteNonQuery();
                        lasignar = lasignar - movto.UnidadesMovto;
                    }
                    if (movto.UnidadesMovto > lasignar)
                    {
                        // dividier 2 movtos el primero movto.unidadesmovto - lasignar y el segundo con lasignar 
                        com.CommandText = "update mgw10010 set ccostoca01 = ccostca01 + " + costounitario + " where cidmovim01 = " + movto.IdMovim;
                        int lafectados = com.ExecuteNonQuery();
                        lasignar = 0;
                    }
                
                }
                    

            }

            
        }

        public void mTraerDataset(List<string> lquery, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen(mEmpresa);
            DataSet ds = new DataSet();
            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            string nombretabla = "Tabla";
            int indice = 1;
            foreach (string lista in lquery)
            {
                OleDbCommand mySqlCommand = new OleDbCommand(lista, lconexion);
                mySqlDataAdapter.SelectCommand = mySqlCommand;
                mySqlDataAdapter.Fill(ds, nombretabla + indice.ToString());
                indice++;
            }
            Datos = ds;
            lconexion.Close();
        }

        public OleDbConnection mAbrirConexionOrigen(string mEmpresa)
        {
            _conexion = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion.Open();
            }
            return _conexion;

        }
    
        
    }
}
