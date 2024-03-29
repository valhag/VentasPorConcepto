﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
using System.Data;
using MyExcel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Data.SqlClient;
using LibreriaDoctos;

namespace VentasPorConcepto
{

    

    public class RegProducto
    {
        public int IdProducto;
        public string CodigoProducto;
        public string NombreProducto;
        public double Existencia;
        public double EntradasPeriodo;
        public double SalidasPeriodo;
        public int metodocosteo;
        public string Clasif1;
        public string Clasif2;
        public string Clasif3;
        public string Clasif4;
        public string Clasif5;
        public string Clasif6;
        
    }

    public class RegCapas
    {
        public int IdProducto;
        public string Fecha;
        public decimal ExistenciaInicial;
        public decimal ExistenciaEntradasPeriodo;
        public decimal ExistenciaSalidasPeriodo;
        public decimal ExistenciaFinal;
        public long idcapae;
        public long idcapas;
        public long idcapa;
        public decimal costo;
        public string  almacen;
    }

    public class RegConcepto
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

    }

    
    public class Class1
    {
        
        ClassConexion lcon = new ClassConexion();

        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public OleDbConnection _conexion;
        private DataTable DatosFacturaAbono = null;

        public DataTable DatosReporte = null;
        private DataTable DatosMaestro = null;
        private DataTable DatosDetalle = null;
        private DataTable DatosEmpresas = null;

        private DataTable DatosReporteAdminpaq = null;

        public List<RegConcepto> _RegClasificaciones = new List<RegConcepto>();

        List<RegProducto> listaprods = new List<RegProducto>();
        List<RegCapas> listacapas = new List<RegCapas>();
        List<RegCapas> sortedlist = new List<RegCapas>();

        public DataSet Datos = null;
        
        public class RegEmpresa
        {
            private string _Nombre;

            public string Nombre
            {
                get { return _Nombre; }
                set { _Nombre = value; }
            }
            private string _Ruta;

            public string Ruta
            {
                get { return _Ruta; }
                set { _Ruta = value; }
            }
        }
        public class RegConcepto
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

        }


        public void mTestFotos()
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];


            string ruta = @"C:\Users\victor\Pictures\Saved Pictures";

            sheet.Rows[lrenglon].RowHeight = 100;
            sheet.Columns["M"].ColumnWidth = 50;

            string lstrPicture = ruta + @"\azael2.jpg";
            //sheet.Cells[5,5].Pictures().Insert(lstrPicture);


            sheet.Rows[lrenglon].RowHeight = 100;
            sheet.Columns["M"].ColumnWidth = 50;

            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[lrenglon, 13];
            Microsoft.Office.Interop.Excel.Range oRange1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[lrenglon+1, 13];
            float Left = (float)((double)oRange.Left);
            float Top = (float)((double)oRange.Top);
            const float ImageSize = 64;

            float height = (float)((double)oRange1.Top - (double)oRange.Top);
            sheet.Shapes.AddPicture(lstrPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left+75, Top,ImageSize, height);

            lrenglon++;
            lstrPicture = ruta + @"\azael3.jpg";


            sheet.Rows[lrenglon].RowHeight = 100;
            sheet.Columns["M"].ColumnWidth = 50;

            oRange = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[lrenglon, 13];
            oRange1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[lrenglon + 1, 13];
            Left = (float)((double)oRange.Left);
            Top = (float)((double)oRange.Top);
            

            height = (float)((double)oRange1.Top - (double)oRange.Top);
            sheet.Shapes.AddPicture(lstrPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left + 75, Top, ImageSize, height);





            /*
            sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOCLIENTE"].ToString().Trim();

            Rows(CStr(lRenglon)).RowHeight = 100
                Columns("M").ColumnWidth = 50


            //configuracionencabezadoPedidoFacturaComercial(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');


                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOCLIENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOAGENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["foliopedido"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                sheet.Cells[lrenglon, lcolumna++].value = row["foliofactura"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //C

                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION2"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION3"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION4"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION5"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION6"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREMONEDA"].ToString().Trim(); //pendiente de facturar

                sheet.Cells[lrenglon, lcolumna++].value = row["CPRECIOCAPTURADO"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CPORCENTAJEDESCUENTO1"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADESCAPTURADAS"].ToString().Trim(); //pendiente de facturar
                sheet.Cells[lrenglon, lcolumna++].value = row["cneto"].ToString().Trim(); //pendiente de facturar


                sheet.Cells[lrenglon, lcolumna++].value = row["CDESCUENTO1"].ToString().Trim(); //pendiente de facturar
                sheet.Cells[lrenglon, lcolumna++].value = row["CIMPUESTO1"].ToString().Trim(); //pendiente de facturar


                sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fechav;

                sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();*/
            return;
        }

        public List<RegEmpresa> mCargarEmpresas(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = mAbrirRutaGlobal(out amensaje);

            List<RegEmpresa> _RegEmpresas = new List<RegEmpresa>();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("select cnombree01,crutadatos from mgw00001 where cidempresa > 1 ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresa lRegEmpresa = new RegEmpresa();
                            lRegEmpresa.Nombre = lreader[0].ToString();
                            lRegEmpresa.Ruta = lreader[1].ToString();
                            _RegEmpresas.Add(lRegEmpresa);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }



            return _RegEmpresas;




        }


        

        public OleDbConnection mAbrirRutaGlobal(out string amensaje)
        {
            amensaje = "";
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIODATOS");
            if (obc == null)
            {
                amensaje = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + obc.ToString();
            try
            {
                _conexion.Open();
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
            }
            return _conexion;

        }
        public MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            excelApp.DisplayAlerts = false;
            return newWorkbook;

        }

        public MyExcel.Workbook mIniciarExcel(string ruta)
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            //MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();

            //MyApp.Visible = false;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Open(@ruta);

            //newWorkbook.Worksheets.Add();
            return newWorkbook;

        }


        public void mTraerDataset(List<string> lquery, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen(mEmpresa);
            DataSet ds = new DataSet();
            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            string nombretabla = "Tabla";
            int indice =1;
            foreach (string lista in lquery)
            {
                OleDbCommand mySqlCommand = new OleDbCommand(lista, lconexion);
                mySqlDataAdapter.SelectCommand = mySqlCommand;
                mySqlDataAdapter.Fill(ds,nombretabla + indice.ToString() );
                indice++;
            }


            //connection.Open();
            //oledbAdapter = new OleDbDataAdapter(firstSql, connection);
            //oledbAdapter.Fill(ds, "First Table");
            //oledbAdapter.SelectCommand.CommandText = secondSql;
            //oledbAdapter.Fill(ds, "Second Table");
            //oledbAdapter.Dispose();
            //connection.Close();

            Datos = ds;

        }



        public RegCliente mValidarCatalogoComercial(int atipo, string aCodigo, string sEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = sEmpresa;

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = Properties.Settings.Default.server;
            string user = Properties.Settings.Default.user;
            string pwd = Properties.Settings.Default.password;
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();


            string lquery = "";
            if (atipo ==1)
                lquery = " select crazonsocial,cidclienteproveedor from admClientes where ccodigocliente = '" + aCodigo + "' and ctipoCLIENTE = 1";
            if (atipo == 2)
                lquery = " select cnombreagente,cidagente from admAgentes where ccodigoagente = '" + aCodigo + "'";
            if (atipo == 3)
                lquery = " select crazonsocial,cidclienteproveedor from admClientes where ccodigocliente = '" + aCodigo + "' and ctipoCLIENTE = 3";

            if (atipo == 4)
                lquery = " select cnombrealmacen,cidalmacen from admAlmacenes where ccodigoalmacen = '" + aCodigo + "'";

            if (atipo == 6)
                lquery = " select cusuario from TablaUsuarios where cusuario = '" + aCodigo + "'";


            SqlCommand mySqlCommand = new SqlCommand(lquery, _conexion1);

            //string lregresa = "";
            SqlDataReader lreader;
            RegCliente lregresa = new RegCliente();
            lreader = mySqlCommand.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                lregresa.RazonSocial = lreader[0].ToString();
                if (atipo != 6)
                {
                    lregresa.Id = long.Parse(lreader[1].ToString());
                }
                
            }
            lreader.Close();
            _conexion1.Close();
            return lregresa;
        }

        public void mTraerDatasetComercial(List<string> lquery, string sEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = sEmpresa;

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = Properties.Settings.Default.server;
            string user = Properties.Settings.Default.user;
            string pwd = Properties.Settings.Default.password;
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();



            DataSet ds = new DataSet();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            string nombretabla = "Tabla";
            int indice = 1;
            foreach (string lista in lquery)
            {
                SqlCommand mySqlCommand = new SqlCommand(lista, _conexion1);
                mySqlDataAdapter.SelectCommand = mySqlCommand;
                mySqlDataAdapter.Fill(ds, nombretabla + indice.ToString());
                indice++;
            }


            //connection.Open();
            //oledbAdapter = new OleDbDataAdapter(firstSql, connection);
            //oledbAdapter.Fill(ds, "First Table");
            //oledbAdapter.SelectCommand.CommandText = secondSql;
            //oledbAdapter.Fill(ds, "Second Table");
            //oledbAdapter.Dispose();
            //connection.Close();

            Datos = ds;

        }


        public void mTraerInformacionPrimerReporte(string lquery, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen(mEmpresa);

            OleDbCommand mySqlCommand = new OleDbCommand(lquery, lconexion);


            DataSet ds = new DataSet();

            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            DatosFacturaAbono = ds.Tables[0];
            DatosReporteAdminpaq = ds.Tables[0];

        }



        public void mTraerInformacionComercial2(string lquery, string mEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = mEmpresa;

            string sempresa ;

            string server = Properties.Settings.Default.server;
            string user = Properties.Settings.Default.user;
            string pwd = Properties.Settings.Default.password;
            sempresa = rutadestino;
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();




            DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, _conexion1);



            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;
            //mySqlDataAdapter.SelectCommand.CommandText = lsql;

            mySqlDataAdapter.Fill(ds);

            DatosMaestro = ds.Tables[0];
            DatosDetalle = ds.Tables[1];
            _conexion1.Close();

            string Cadenaconexion2 = "data source =" + server + ";initial catalog = 'CompacWAdmin';user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion2;
            _conexion1.Open();




            DataSet ds1 = new DataSet();

            string lsql2 = "SELECT  [CIDEMPRESA] " + 
      " ,[CNOMBREEMPRESA] " + 
  " FROM [Empresas]";
            SqlDataAdapter mySqlDataAdapter1 = new SqlDataAdapter(lsql2, _conexion1);
            mySqlDataAdapter1.Fill(ds1);

            DatosEmpresas = ds1.Tables[0];

           

        }



        public void mTraerInformacionClasificacionesComercial(StringBuilder lquery, string mEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
                //            rutadestino = "c:\\compacw\\empresas\\adtala2";
                string rutadestino = mEmpresa;

                string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

                string server = Properties.Settings.Default.server;
                string user = Properties.Settings.Default.user;
                string pwd = Properties.Settings.Default.password;
                //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
                //string lruta3 = obc.ToString();
                string lruta4 = @rutadestino;
                _conexion1 = new SqlConnection();
                string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
                _conexion1.ConnectionString = Cadenaconexion1;
                _conexion1.Open();




                DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql,_conexion1);
            
            

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;
            //mySqlDataAdapter.SelectCommand.CommandText = lsql;

            mySqlDataAdapter.Fill(ds);

            DataTable DatosClasif = ds.Tables[0];
            _RegClasificaciones.Clear();
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                RegConcepto clasif = new RegConcepto();
                clasif.id = long.Parse(row["CIDVALORCLASIFICACION"].ToString());
                clasif.Nombre = row["CVALORCLASIFICACION"].ToString();
                _RegClasificaciones.Add(clasif);
            }

            _conexion1.Close();

        }


        public void mTraerInformacionComercial(StringBuilder lquery, string mEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
                //            rutadestino = "c:\\compacw\\empresas\\adtala2";
                string rutadestino = mEmpresa;

                string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

                string server = Properties.Settings.Default.server;
                string user = Properties.Settings.Default.user;
                string pwd = Properties.Settings.Default.password;
                //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
                //string lruta3 = obc.ToString();
                string lruta4 = @rutadestino;
                _conexion1 = new SqlConnection();
                string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
                _conexion1.ConnectionString = Cadenaconexion1;
                _conexion1.Open();




                DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql,_conexion1);
            
            

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;
            //mySqlDataAdapter.SelectCommand.CommandText = lsql;

            mySqlDataAdapter.Fill(ds);

            DatosReporte = ds.Tables[0];
            if (ds.Tables.Count > 1)
                DatosDetalle = ds.Tables[1];
            _conexion1.Close();

        }

        

        private void TotalizarReporteFacturaAbono(MyExcel.Worksheet sheet, int lrenglon)
        {
            int lrengloninicial = 7;
            int lrenglonfinal = lrenglon-1;
            

            //sheet.Cells[lrenglon, 2].value = "Totales";
            sheet.get_Range("I" + lrenglon.ToString(), "I" +
            lrenglon.ToString()).Formula = "=SUM(I" + lrengloninicial.ToString() + ":I" + lrenglonfinal.ToString() + ")";
            sheet.get_Range("J" + lrenglon.ToString(), "J" +
            lrenglon.ToString()).Formula = "=SUM(J" + lrengloninicial.ToString() + ":J" + lrenglonfinal.ToString() + ")";
            sheet.get_Range("K" + lrenglon.ToString(), "K" +
            lrenglon.ToString()).Formula = "=SUM(K" + lrengloninicial.ToString() + ":K" + lrenglonfinal.ToString() + ")";
            sheet.get_Range("L" + lrenglon.ToString(), "L" +
            lrenglon.ToString()).Formula = "=SUM(L" + lrengloninicial.ToString() + ":L" + lrenglonfinal.ToString() + ")";
            sheet.get_Range("M" + lrenglon.ToString(), "M" +
            lrenglon.ToString()).Formula = "=SUM(M" + lrengloninicial.ToString() + ":M" + lrenglonfinal.ToString() + ")";

            sheet.get_Range("N" + lrenglon.ToString(), "N" +
            lrenglon.ToString()).Formula = "=SUM(N" + lrengloninicial.ToString() + ":N" + lrenglonfinal.ToString() + ")";

            
            sheet.get_Range("I" + lrenglon.ToString(), "N" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("I" + lrenglon, "N" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("I" + lrenglon, "N" + lrenglon).Font.Bold = true;

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

        private void EncabezadoEmpresa(MyExcel.Worksheet sheet,string Empresa, string texto)
        {
            sheet.Cells[2, 3].value = Empresa;
            sheet.get_Range("C2").Font.Size = 16;
            sheet.get_Range("C2").Font.Bold = true;
            sheet.get_Range("C2", "D2").Merge();
            sheet.Cells[3, 3].value = texto;
            sheet.get_Range("C3", "E3").Merge();
            sheet.get_Range("C3").Font.Bold = true;
        }

        private void configuracionencabezadoPedidoFactura(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, string lfecha1, string lfecha2)
        {
            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";

            sheet.Cells[1, 7].value = lfecha1;
            sheet.Cells[2, 7].value = lfecha2;
            


//            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


            sheet.Cells[lrenglon, lcolumna++].value = "Fecha";
            sheet.Cells[lrenglon, lcolumna++].value = "Serie pedido";
            sheet.Cells[lrenglon, lcolumna++].value = "# pedidos";
            sheet.Cells[lrenglon, lcolumna++].value = "cliente";
            sheet.Cells[lrenglon, lcolumna++].value = "importe";
            sheet.Cells[lrenglon, lcolumna++].value = "pendiente de facturar";

            sheet.Cells[lrenglon, lcolumna++].value = "Fecha de Facturacion";

            sheet.Cells[lrenglon, lcolumna++].value = "Serie factura";
            sheet.Cells[lrenglon, lcolumna++].value = "# de factura";
            sheet.Cells[lrenglon, lcolumna++].value = "cliente";
            sheet.Cells[lrenglon, lcolumna++].value = "importe";
            sheet.Cells[lrenglon, lcolumna++].value = "Impuesto";
            sheet.Cells[lrenglon, lcolumna++].value = "Retención";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";
        }

        private void configuracionencabezadooc(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, string lfecha1, string lfecha2)
        {
            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";

            sheet.Cells[1, 7].value = lfecha1;
            sheet.Cells[2, 7].value = lfecha2;



            //            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


            sheet.Cells[lrenglon, lcolumna++].value = "Prog.";
            sheet.Cells[lrenglon, lcolumna++].value = "Fecha";
            sheet.Cells[lrenglon, lcolumna++].value = "Folio";
            sheet.Cells[lrenglon, lcolumna++].value = "Proveedor";
            sheet.Cells[lrenglon, lcolumna++].value = "Codigo";
            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Alterno";
            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre Alterno";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad Solicitada";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad Pendiente";

        }




        private void configuracionencabezadocapas(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, string lfecha1, string lfecha2)
        {
            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";

            sheet.Cells[1, 7].value = lfecha1;
            sheet.Cells[2, 7].value = lfecha2;

            sheet.get_Range("A" + lrenglon.ToString(), "T" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "T" +
            lrenglon.ToString()).Font.Color = Color.White;

            //            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


            		//		TIPO	Inventario Inicial	Entradas	Salidas	Existencia	Inventario Inicial	Entradas	Salidas	Inventario Final	COSTO DE LA CAPA


            sheet.Cells[lrenglon, lcolumna++].value = "Tipo Movto.";
            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre";
            sheet.Cells[lrenglon, lcolumna++].value = "Método Costeo";
            sheet.Cells[lrenglon, lcolumna++].value = "Inventario Inicial";
            sheet.Cells[lrenglon, lcolumna++].value = "Entradas";
            sheet.Cells[lrenglon, lcolumna++].value = "Salidas";
            sheet.Cells[lrenglon, lcolumna++].value = "Existencia";
            sheet.Cells[lrenglon, lcolumna++].value = "Inventario Inicial";
            sheet.Cells[lrenglon, lcolumna++].value = "Entradas";
            sheet.Cells[lrenglon, lcolumna++].value = "Salidas";
            sheet.Cells[lrenglon, lcolumna++].value = "Inventario Final";
            sheet.Cells[lrenglon, lcolumna++].value = "Costo Capa";
            sheet.Cells[lrenglon, lcolumna++].value = "Almacen";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 1";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 2";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 3";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 4";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 5";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasif 6";



        }

        private void configuracionencabezadoieps(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, string lfecha1, string lfecha2)
        {
            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";

            sheet.Cells[1, 7].value = lfecha1;
            sheet.Cells[2, 7].value = lfecha2;



            //            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


            sheet.Cells[lrenglon, lcolumna++].value = "Foliocargo";
            sheet.Cells[lrenglon, lcolumna++].value = "Seriecargo";
            sheet.Cells[lrenglon, lcolumna++].value = "Fechacargo";
            sheet.Cells[lrenglon, lcolumna++].value = "Cliente";
            sheet.Cells[lrenglon, lcolumna++].value = "Conceptocargo";
            sheet.Cells[lrenglon, lcolumna++].value = "Netocargo";
            sheet.Cells[lrenglon, lcolumna++].value = "IVA";
            sheet.Cells[lrenglon, lcolumna++].value = "IEPS";
            sheet.Cells[lrenglon, lcolumna++].value = "Total Cargo";
            sheet.Cells[lrenglon, lcolumna++].value = "Folio Abono";
            sheet.Cells[lrenglon, lcolumna++].value = "Serie Abono";
            sheet.Cells[lrenglon, lcolumna++].value = "Fecha Abono";
            sheet.Cells[lrenglon, lcolumna++].value = "Concepto Abono";
            sheet.Cells[lrenglon, lcolumna++].value = "Referencia";
            sheet.Cells[lrenglon, lcolumna++].value = "Observaciones";
            sheet.Cells[lrenglon, lcolumna++].value = "Total Abono";
            sheet.Cells[lrenglon, lcolumna++].value = "Total Pagado";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto Pagado";
            sheet.Cells[lrenglon, lcolumna++].value = "Iva Pagado";
            sheet.Cells[lrenglon, lcolumna++].value = "IEPS Pagado";



        }



        private void configuracionencabezadofotos(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, string lfecha1, string lfecha2)
        {
            //EncabezadoEmpresa(sheet, Empresa, texto);
            sheet.Cells[3, 13].value = Empresa.Trim() ;
            sheet.get_Range("M3").Font.Bold = true;
            sheet.Cells[3, 13].HorizontalAlignment = MyExcel.XlHAlign.xlHAlignCenter;


            sheet.Cells[5, 13].value = "MERCANCIA EN PRODUCCION";
            sheet.get_Range("M5").Font.Bold = true;
            sheet.Cells[5, 13].HorizontalAlignment = MyExcel.XlHAlign.xlHAlignCenter;
            sheet.Cells[6, 13].value = "ORDENADO POR FECHA";
            sheet.get_Range("M6").Font.Bold = true;
            sheet.Cells[6, 13].HorizontalAlignment = MyExcel.XlHAlign.xlHAlignCenter;


            sheet.Cells[2, 31].value = "Fecha :" + DateTime.Today.ToShortDateString();



            int lcolumna = 1;

            //sheet.Cells[1, 6].value = "Fecha Inicial";
            //sheet.Cells[2, 6].value = "Fecha Final";

           

        }
        

        private void configuracionencabezadoFacturaAbono(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon)
        {
            sheet.Cells[2, 3].value = Empresa;
            sheet.get_Range("C2").Font.Size = 16;
            sheet.get_Range("C2").Font.Bold = true;
            sheet.get_Range("C2", "D2").Merge();
            sheet.Cells[3, 3].value = texto;
            sheet.get_Range("C3", "E3").Merge();
            sheet.get_Range("C3").Font.Bold = true;

            sheet.get_Range("A" + lrenglon.ToString(), "P" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "P" +
            lrenglon.ToString()).Font.Color = Color.White;


/*            Concepto	Folio	Serie	Importe	Referencia	Saldo		Serie	Folio	"Pago en 
Efectivo"	"Pago TC
Bancomer"	"Pago TC
IXE"	TRANSFER.	"Pago con 
Cheque"	Devoluciones
	*/
            int lcolumna = 1;

            sheet.Cells[lrenglon, lcolumna++].value = "Concepto";
            sheet.Cells[lrenglon, lcolumna++].value = "Folio";
            sheet.Cells[lrenglon, lcolumna++].value = "Serie";
            sheet.Cells[lrenglon, lcolumna++].value = "Importe";
            sheet.Cells[lrenglon, lcolumna++].value = "Referencia";
            sheet.Cells[lrenglon, lcolumna++].value = "Saldo";
            sheet.Cells[lrenglon, lcolumna++].value = "Serie";
            sheet.Cells[lrenglon, lcolumna++].value = "Folio";
            sheet.Cells[lrenglon, lcolumna++].value = "Pago en Efectivo";
            sheet.Cells[lrenglon, lcolumna++].value = "Pago TC Bancomer";
            sheet.Cells[lrenglon, lcolumna++].value = "Pago TC IXE";
            sheet.Cells[lrenglon, lcolumna++].value = "TRANSFER";
            sheet.Cells[lrenglon, lcolumna++].value = "Pago con Cheque";
            sheet.Cells[lrenglon, lcolumna++].value = "NC";
            sheet.Cells[lrenglon, lcolumna++].value = "Devoluciones";


            



        }


        private void mConfigurarObjetosSQLImpresion()
        {
            DataTable Productos = Datos.Tables[0];
            DataTable InventarioInicialentradas = Datos.Tables[1];
            DataTable InventarioInicialsalidas = Datos.Tables[2];
            DataTable capasinicialentradas = Datos.Tables[3];
            DataTable capasinicialsalidas = Datos.Tables[4];
            DataTable Movimientosentradas = Datos.Tables[5];
            DataTable Movimientossalidas = Datos.Tables[6];
            DataTable capasenperiodoentradas = Datos.Tables[7];
            DataTable capasenperiodosalidas = Datos.Tables[8];

            listaprods.Clear();
            listacapas.Clear();

            DataRow xxx;
            DataTable table1 = new DataTable();
            table1.Columns.Add("uno", typeof(int));
            table1.Columns.Add("dos", typeof(double));
            table1.Columns.Add("tres", typeof(double));
            table1.Columns.Add("cuatro", typeof(double));
            table1.Columns.Add("cinco", typeof(double));
            table1.Columns.Add("seis", typeof(double));
            table1.Columns.Add("siete", typeof(double));
            xxx = table1.Rows.Add(0, 0, 0, 0, 0, 0, 0);

            var inventarioinicial = from p in Productos.AsEnumerable()
                                    //from e in InventarioInicialentradas.AsEnumerable()
                                    join e in InventarioInicialentradas.AsEnumerable() on (string)p["idprodue"].ToString() equals (string)e["idprodue"].ToString() into tempp
                                    from e1 in tempp.DefaultIfEmpty(xxx)
                                    join s in InventarioInicialsalidas.AsEnumerable() on (string)e1[0].ToString() equals (string)s["idprodus"].ToString() into temp
                                    from s1 in temp.DefaultIfEmpty(xxx)
                                    join me in Movimientosentradas.AsEnumerable() on (string)p["idprodue"].ToString() equals (string)me["idprodue"].ToString() into temp1
                                    from move in temp1.DefaultIfEmpty(xxx)
                                    join ms in Movimientossalidas.AsEnumerable() on (string)p["idprodue"].ToString() equals (string)ms["idprodus"].ToString() into temp2
                                    from movs in temp2.DefaultIfEmpty(xxx)
                                    select new
                                    {
                                        Id = p.Field<int>(0).ToString(),
                                        Nombre = p.Field<string>(2).ToString(),
                                        Salidas = s1.Field<double>(4).ToString() ?? string.Empty,
                                        Entradas = e1.Field<double>(1),
                                        Codigo = p.Field<string>(1).ToString(),
                                        Metodo = p.Field<int>(3).ToString(),
                                        MovEntradas = move.Field<double>(1),
                                        MovSalidas = movs.Field<double>(1),
                                        clasif1 = p.Field<string>(4).ToString(),
                                        clasif2 = p.Field<string>(5).ToString(),
                                        clasif3 = p.Field<string>(6).ToString(),
                                        clasif4 = p.Field<string>(7).ToString(),
                                        clasif5 = p.Field<string>(8).ToString(),
                                        clasif6 = p.Field<string>(9).ToString()
                                    };
            //Codigo = e.Field<string>(1).ToString()
            //Metodo = e.Field<int>(3).ToString()
            string nombre = "";
            double saldo = 0;
            int cuantosprods = 0;
            foreach (var saldos in inventarioinicial)
            {
                RegProducto lprod = new RegProducto();
                nombre = saldos.Nombre;
                lprod.IdProducto = int.Parse(saldos.Id);
                lprod.NombreProducto = saldos.Nombre;
                lprod.CodigoProducto = saldos.Codigo;
                lprod.metodocosteo = int.Parse(saldos.Metodo);
                saldo = saldos.Entradas - double.Parse(saldos.Salidas);
                lprod.Existencia = saldo;
                lprod.EntradasPeriodo = saldos.MovEntradas;
                lprod.SalidasPeriodo = saldos.MovSalidas;
                lprod.Clasif1 = saldos.clasif1;
                lprod.Clasif2 = saldos.clasif2;
                lprod.Clasif3 = saldos.clasif3;
                lprod.Clasif4 = saldos.clasif4;
                lprod.Clasif5 = saldos.clasif5;
                lprod.Clasif6 = saldos.clasif6;
                listaprods.Add(lprod);

            }

            //and (string)e["cidcapa"].ToString() equals (string)s["cidcapa"].ToString() into UP

            DataRow zz;
            DataTable table = new DataTable();
            table.Columns.Add("uno", typeof(double));
            table.Columns.Add("dos", typeof(double));
            table.Columns.Add("tres", typeof(double));
            table.Columns.Add("cuatro", typeof(double));
            table.Columns.Add("cinco", typeof(double));
            table.Columns.Add("seis", typeof(double));
            table.Columns.Add("siete", typeof(double));
            table.Columns.Add("ocho", typeof(double));
            table.Columns.Add("nueve", typeof(double));
            table.Columns.Add("diez", typeof(double));
            table.Columns.Add("once", typeof(double));
            zz = table.Rows.Add(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            var capasinicial = from capaseinicial in capasinicialentradas.AsEnumerable()
                               join capassinicial in capasinicialsalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidproducto"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassinicial["cidproducto"].ToString(),
                                   cidcapa = capassinicial["cidcapa"].ToString()
                               } into temp
                               from s1 in temp.DefaultIfEmpty(zz)
                               join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidproducto"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidproducto"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                               from s2 in temp2.DefaultIfEmpty(zz)
                               /*join capasentradasperiodo in capasenperiodoentradas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   //cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                                * 
                               equals
                               new
                               {
                                   cidprodu01 = capasentradasperiodo["cidprodu01"].ToString(),
                                   //cidcapa = capasentradasperiodo["cidcapa"].ToString()
                               } into temp3
                               from s3 in temp3.DefaultIfEmpty(zz)*/
                               select new
                               {
                                   cidprodu = capaseinicial.Field<int>(0),
                                   cidcapa = capaseinicial.Field<int>(1),
                                   cfecha = capaseinicial.Field<string>(2),
                                   unidadesentrada = capaseinicial.Field<double>(3),
                                   unidadessalida = s1.Field<double>(2).ToString() ?? string.Empty,
                                   unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                   costo = capaseinicial.Field<double>(4),
                                   almacen = capaseinicial.Field<string>(5)
                               };
            decimal existenciacapae = 0;
            decimal existenciacapas = 0;

            foreach (var capa in capasinicial)
            {
                RegCapas capalocal = new RegCapas();
                capalocal.Fecha = capa.cfecha;
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;

                existenciacapas = decimal.Parse(capa.unidadessalida.ToString()); //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                //capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada);
                //if (capalocal.ExistenciaEntradasPeriodo > 0)
                // capalocal.ExistenciaInicial = 0;
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                //capalocal.idcapae = long.Parse(capa.idcapae);
                //capalocal.idcapas = long.Parse(capa.idcapas);
                capalocal.idcapa = long.Parse(capa.cidcapa.ToString());
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                if (existenciacapae - existenciacapas != 0)
                    listacapas.Add(capalocal);
            }
            
            var capassoloentrdas = from capaseinicial in capasenperiodoentradas.AsEnumerable()
                                   join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidproducto"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidproducto"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                                   from s2 in temp2.DefaultIfEmpty(zz)
                                   select new
                                   {
                                       //cidprodu = capaseinicial.Field<decimal>(0),
                                       cidprodu = capaseinicial.Field<int>(0),
                                       cidcapa = capaseinicial.Field<int>(1),
                                       unidadesentrada = 0,
                                       unidadesperiodoentrada = capaseinicial.Field<double>(2),
                                       costo = capaseinicial.Field<double>(3),
                                       unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                       almacen = capaseinicial.Field<string>(4)
                                   };
            foreach (var capa in capassoloentrdas)
            {
                RegCapas capalocal = new RegCapas();
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;
                existenciacapas = 0; //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada.ToString());
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                listacapas.Add(capalocal);
            }
            listacapas = listacapas.OrderBy(o => o.IdProducto).ToList();
            //sortedlist.Clear();
            //sortedlist = listacapas.OrderBy(o => o.IdProducto).ToList();
            //listacapas.Clear();
            //listacapas = sortedlist;

        }

        private void mConfigurarObjetosImpresion()
        {
            DataTable Productos = Datos.Tables[0];
            DataTable InventarioInicialentradas = Datos.Tables[1];
            DataTable InventarioInicialsalidas = Datos.Tables[2];
            DataTable capasinicialentradas = Datos.Tables[3];
            DataTable capasinicialsalidas = Datos.Tables[4];
            DataTable Movimientosentradas = Datos.Tables[5];
            DataTable Movimientossalidas = Datos.Tables[6];
            DataTable capasenperiodoentradas = Datos.Tables[7];
            DataTable capasenperiodosalidas = Datos.Tables[8];

            listaprods.Clear();
            listacapas.Clear();

            DataRow xxx;
            DataTable table1 = new DataTable();
            table1.Columns.Add("uno", typeof(double));
            table1.Columns.Add("dos", typeof(double));
            table1.Columns.Add("tres", typeof(double));
            table1.Columns.Add("cuatro", typeof(double));
            table1.Columns.Add("cinco", typeof(double));
            table1.Columns.Add("seis", typeof(double));
            table1.Columns.Add("siete", typeof(double));
            xxx = table1.Rows.Add(0, 0, 0,0,0,0,0);

            var inventarioinicial = from p in Productos.AsEnumerable()
                                    //from e in InventarioInicialentradas.AsEnumerable()
                                    join e in InventarioInicialentradas.AsEnumerable() on (string)p["idprodue"].ToString() equals (string)e["idprodue"].ToString() into tempp
                                    from e1 in tempp.DefaultIfEmpty(xxx)
                                    join s in InventarioInicialsalidas.AsEnumerable() on (string)e1[0].ToString() equals (string)s["idprodus"].ToString() into temp
                                    from s1 in temp.DefaultIfEmpty(xxx)
                                    join me in Movimientosentradas.AsEnumerable() on (string)s1[0].ToString() equals (string)me["idprodue"].ToString() into temp1
                                    from move in temp1.DefaultIfEmpty(xxx)
                                    join ms in Movimientossalidas.AsEnumerable() on (string)s1[0].ToString() equals (string)ms["idprodus"].ToString() into temp2
                                    from movs in temp2.DefaultIfEmpty(xxx)
                                    select new
                                    {
                                        //Id = e.Field<decimal>(0).ToString(),
                                        Id = p.Field<decimal>(0).ToString(),
                                        //Nombre = e.Field<string>(2).ToString(),
                                        Nombre = p.Field<string>(2).ToString(),
                                        Salidas = s1.Field<double>(4).ToString() ?? string.Empty,
                                        //Entradas = e.Field<double>(4),
                                        Entradas = e1.Field<double>(1),
                                        Codigo = p.Field<string>(1).ToString(),
                                        Metodo = p.Field<decimal>(3).ToString(),
                                        MovEntradas = move.Field<double>(1),
                                        MovSalidas = movs.Field<double>(1),
                                        clasif1 = p.Field<string>(4).ToString(),
                                        clasif2 = p.Field<string>(5).ToString(),
                                        clasif3 = p.Field<string>(6).ToString(),
                                        clasif4 = p.Field<string>(7).ToString(),
                                        clasif5 = p.Field<string>(8).ToString(),
                                        clasif6 = p.Field<string>(9).ToString()
                                    };
            //Codigo = e.Field<string>(1).ToString()
            //Metodo = e.Field<int>(3).ToString()
            string nombre = "";
            double saldo = 0;
            int cuantosprods = 0;
            foreach (var saldos in inventarioinicial)
            {
                RegProducto lprod = new RegProducto();
                nombre = saldos.Nombre;
                lprod.IdProducto = int.Parse(saldos.Id);
                lprod.NombreProducto = saldos.Nombre;
                lprod.CodigoProducto = saldos.Codigo;
                lprod.metodocosteo = int.Parse(saldos.Metodo);
                saldo = saldos.Entradas - double.Parse(saldos.Salidas);
                lprod.Existencia = saldo;
                lprod.EntradasPeriodo = saldos.MovEntradas;
                lprod.SalidasPeriodo = saldos.MovSalidas;
                lprod.Clasif1 = saldos.clasif1;
                lprod.Clasif2 = saldos.clasif2;
                lprod.Clasif3 = saldos.clasif3;
                lprod.Clasif4 = saldos.clasif4;
                lprod.Clasif5 = saldos.clasif5;
                lprod.Clasif6 = saldos.clasif6;
                listaprods.Add(lprod);

            }

            //and (string)e["cidcapa"].ToString() equals (string)s["cidcapa"].ToString() into UP

            DataRow zz;
            DataTable table = new DataTable();
            table.Columns.Add("uno", typeof(double));
            table.Columns.Add("dos", typeof(double));
            table.Columns.Add("tres", typeof(double));
            table.Columns.Add("cuatro", typeof(double));
            table.Columns.Add("cinco", typeof(double));
            table.Columns.Add("seis", typeof(double));
            table.Columns.Add("siete", typeof(double));
            table.Columns.Add("ocho", typeof(double));
            table.Columns.Add("nueve", typeof(double));
            table.Columns.Add("diez", typeof(double));
            table.Columns.Add("once", typeof(double));
            zz = table.Rows.Add(0, 0, 0, 0, 0, 0,0, 0, 0,0, 0);

            var capasinicial = from capaseinicial in capasinicialentradas.AsEnumerable()
                               join capassinicial in capasinicialsalidas.AsEnumerable() on
                               new 
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassinicial["cidprodu01"].ToString(),
                                   cidcapa = capassinicial["cidcapa"].ToString()
                               } into temp
                               from s1 in temp.DefaultIfEmpty(zz)
                               join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidprodu01"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                               from s2 in temp2.DefaultIfEmpty(zz)
                               /*join capasentradasperiodo in capasenperiodoentradas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   //cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                                * 
                               equals
                               new
                               {
                                   cidprodu01 = capasentradasperiodo["cidprodu01"].ToString(),
                                   //cidcapa = capasentradasperiodo["cidcapa"].ToString()
                               } into temp3
                               from s3 in temp3.DefaultIfEmpty(zz)*/
                               select new
                               {
                                   cidprodu = capaseinicial.Field<decimal>(0),
                                   cidcapa = capaseinicial.Field<decimal>(1),
                                   cfecha = capaseinicial.Field<string>(2),
                                   unidadesentrada = capaseinicial.Field<double>(3),
                                   unidadessalida = s1.Field<double>(2).ToString() ?? string.Empty,
                                   unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                   costo = capaseinicial.Field<double>(4),
                                   almacen = capaseinicial.Field<string>(5)
                               };
            decimal existenciacapae = 0;
            decimal existenciacapas = 0;

            foreach (var capa in capasinicial)
            {
                RegCapas capalocal = new RegCapas();
                capalocal.Fecha = capa.cfecha;
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;

                existenciacapas = decimal.Parse(capa.unidadessalida.ToString()); //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                //capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada);
                //if (capalocal.ExistenciaEntradasPeriodo > 0)
                   // capalocal.ExistenciaInicial = 0;
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                //capalocal.idcapae = long.Parse(capa.idcapae);
                //capalocal.idcapas = long.Parse(capa.idcapas);
                capalocal.idcapa = long.Parse(capa.cidcapa.ToString());
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                if (existenciacapae - existenciacapas != 0)
                        listacapas.Add(capalocal);
            }
            var capassoloentrdas = from capaseinicial in capasenperiodoentradas.AsEnumerable()
                                   join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidprodu01"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                                   from s2 in temp2.DefaultIfEmpty(zz)
                                   select new
                                   {
                                       cidprodu = capaseinicial.Field<decimal>(0),
                                       cidcapa = capaseinicial.Field<decimal>(1),
                                       //cfecha = capaseinicial.Field<string>(2),
                                       unidadesentrada = 0,
                                       unidadesperiodoentrada = capaseinicial.Field<double>(2),
                                       costo = capaseinicial.Field<double>(3),
                                       unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                       almacen = capaseinicial.Field<string>(4)
                                   };
            foreach (var capa in capassoloentrdas)
            {
                RegCapas capalocal = new RegCapas();
                //capalocal.Fecha = capa.cfecha;
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;
                existenciacapas = 0; //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada.ToString());
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                    listacapas.Add(capalocal);
            }
            listacapas = listacapas.OrderBy(o => o.IdProducto).ToList();
            //sortedlist.Clear();
            //sortedlist = listacapas.OrderBy(o => o.IdProducto).ToList();
            //listacapas.Clear();
            //listacapas = sortedlist;

        }

        //mReporteReporteComisiones
        public void mReporteReporteComisiones()
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            //configuracionencabezadoComisiones(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);
            //MyExcel.DisplayAlerts = False 

            lrenglon = configuracionencabezadoComisionesRemision(sheet);
            string lultimocliente = "";
            decimal ltotal1 = 0;
            double ltotal2 = 0;

            decimal ltotal015 = 0;
            decimal ltotal1530 = 0;
            decimal ltotal3045 = 0;
            decimal ltotal4560 = 0;
            decimal ltotal60 = 0;
            decimal ltotal015a = 0;
            decimal ltotal1530a = 0;
            decimal ltotal3045a = 0;
            decimal ltotal4560a = 0;
            decimal ltotal60a = 0;

            long lultimofolio = 0;

            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                int lcolumna = 1;


                if (row["CCODIGOCLIENTE"].ToString().Trim() != lultimocliente && lultimocliente != "")
                {
                   /* sheet.Cells[lrenglon, 10].value = ltotal1;
                    sheet.Cells[lrenglon, 16].value = ltotal2;

                    sheet.Cells[lrenglon, 19].value = ltotal015;
                    sheet.Cells[lrenglon, 20].value = ltotal1530;
                    sheet.Cells[lrenglon, 21].value = ltotal3045;
                    sheet.Cells[lrenglon, 22].value = ltotal4560;
                    sheet.Cells[lrenglon, 23].value = ltotal60;

                    sheet.Cells[lrenglon, 25].value = ltotal015a;
                    sheet.Cells[lrenglon, 26].value = ltotal1530a;
                    sheet.Cells[lrenglon, 27].value = ltotal3045a;
                    sheet.Cells[lrenglon, 28].value = ltotal4560a;
                    sheet.Cells[lrenglon, 29].value = ltotal60a;
                    */
                    //lrenglon += 3;

                    /*
                    ltotal1 = 0;
                    ltotal2 = 0;
                    ltotal015 = 0;
                    ltotal1530 = 0;
                    ltotal3045 = 0;
                    ltotal4560 = 0;
                    ltotal60 = 0;

                    ltotal015a = 0;
                    ltotal1530a = 0;
                    ltotal3045a = 0;
                    ltotal4560a = 0;
                    ltotal60a = 0;*/
                }
                //else
                //  ltotal1 += decimal.Parse(row["CPENDIENTE"].ToString().Trim());

<<<<<<< HEAD
                if (lultimofolio != long.Parse(row["CFOLIO"].ToString().Trim()))
                {

                    //lrenglon++;
                    if (ltotal1 > 0)
                    {
                        sheet.Cells[lrenglon, 12].value = ltotal1;
                        sheet.Cells[lrenglon, 15].value = ltotal2;
                        lrenglon += 2;
                        ltotal1 = 0;
                        ltotal2 = 0;
                    }

                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOCLIENTE"].ToString().Trim();
                    lultimocliente = row["CCODIGOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo

                    //sheet.Cells[lrenglon, lcolumna++].value = row["CDIASCREDITOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CSERIEDOCUMENTO"].ToString().Trim(); //Fecha Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Fecha Cargo
                    string sFecha1 = "'" + row["cfechad"].ToString().Trim();
                    sFecha1 = sFecha1.Substring(0, 11);
                    sheet.Cells[lrenglon, lcolumna++].value = sFecha1;
                    sFecha1 = "'" + row["cfechap"].ToString().Trim();
                    /*if (sFecha1 != "'")
                    {
                        sFecha1 = sFecha1.Substring(0, 11);
                        sheet.Cells[lrenglon, lcolumna++].value = sFecha1; //Fecha Cargo
                    }
                    else
                        lcolumna++;*/
                    sheet.Cells[lrenglon, lcolumna++].value = row["CTOTAL"].ToString().Trim(); //Fecha Cargo

                    //sheet.Cells[lrenglon, lcolumna++].value = row["CPENDIENTE"].ToString().Trim(); //Fecha Cargo

                    

                    //sheet.Cells[lrenglon, lcolumna++].value = row["numdias"].ToString().Trim(); //Fecha Cargo
                    lultimofolio = long.Parse(row["CFOLIO"].ToString().Trim());

                }
                else
                    lcolumna += 7;

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADESCAPTURADAS"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CABREVIATURA"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim(); //Fecha Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["ctotalmov"].ToString().Trim(); //Fecha Cargo

                ltotal1 += decimal.Parse(row["ctotalmov"].ToString().Trim());


                sheet.Cells[lrenglon, lcolumna++].value = row["comision"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                ltotal2 += double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);

                /*


                lcolumna = 19;
                if (int.Parse(row["numdias"].ToString().Trim()) <= 15 && int.Parse(row["numdias"].ToString().Trim()) > 0)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal015 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }

                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 15 && int.Parse(row["numdias"].ToString().Trim()) <= 30)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal1530 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);

                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 30 && int.Parse(row["numdias"].ToString().Trim()) <= 45)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal3045 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 45 && int.Parse(row["numdias"].ToString().Trim()) <= 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal4560 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal60 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }
                lcolumna++;
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) <= 15 && int.Parse(row["numdias"].ToString().Trim()) > 0)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal015a += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 15 && int.Parse(row["numdias"].ToString().Trim()) <= 30)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal1530a += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);

                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 30 && int.Parse(row["numdias"].ToString().Trim()) <= 45)
                {
                    sheet.Cells[lrenglon, lcolumna].value = (double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100)) * 0.75;
                    ltotal3045a += (decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100)) * (decimal)0.75;
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 45 && int.Parse(row["numdias"].ToString().Trim()) <= 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = (double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100)) * 0.5;
                    ltotal4560a += (decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100)) * (decimal)0.5;

                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = 0;
                    ltotal60a += 0;
                }

                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                //sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";



                ltotal2 += decimal.Parse(row["ctotalmov"].ToString().Trim());*/
=======
                //sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADES"].ToString().Trim();
                //sheet.Cells[lrenglon, lcolumna].numberformat = "0.00";
                sheet.Cells[lrenglon, lcolumna++].value = row["CNETO"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CTOTAL"].ToString().Trim(); //Fecha Cargo
                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                
                sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).Style = "Comma";
                sheet.get_Range("D" + lrenglon.ToString(), "E" + lrenglon.ToString()).Style = "Currency";
>>>>>>> 9a41ea45bd8e9002eb6a577c27983ff67c519b3f

                lrenglon++;


            }

            sheet.get_Range("G7", "G" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("L7", "L" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("O7", "O" + lrenglon.ToString()).Style = "Currency";

        }



<<<<<<< HEAD
        public void mReporteComisiones()
=======
        public void mReporteForrajeraComercial(string mEmpresa, DateTime lfechai, DateTime lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoForrajeraComercial(sheet, mEmpresa, "REPORTE DE MOVIMIENTOS POR CONCEPTO POR PRODUCTO", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon = 6;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            string conceptoprevio="" ;
            decimal lcantidad = 0;
            decimal lcosto = 0;
            decimal ltotal = 0;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                string concepto = row["CNOMBRECONCEPTO"].ToString().Trim();
                if (concepto != conceptoprevio)
                {
                    if (conceptoprevio != "")
                    { 
                        // totales

                        sheet.Cells[lrenglon, 1].value = "Total del Concepto";
                        sheet.Cells[lrenglon, 3].value = lcantidad;
                        //sheet.Cells[lrenglon, lcolumna].numberformat = "0.00";
                        sheet.Cells[lrenglon, 4].value = lcosto; //Serie Cargo
                        sheet.Cells[lrenglon, 5].value = ltotal; //Fecha Cargo
                        sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).Style = "Comma";
                        sheet.get_Range("D" + lrenglon.ToString(), "E" + lrenglon.ToString()).Style = "Currency";
                        sheet.get_Range("A" + lrenglon.ToString(), "E" + lrenglon.ToString()).Font.Bold = true;
                        lrenglon += 2;

                    }
                    // imprimir titulo de concepto
                    sheet.Cells[++lrenglon, 1].value = "Concepto: " + concepto;

                    sheet.get_Range("A" + lrenglon.ToString(), "A" + lrenglon.ToString()).Font.Bold = true;


                    
                    lrenglon +=2;
                    conceptoprevio = concepto;
                    lcantidad = 0;
                    lcosto = 0;
                    ltotal = 0;
                }
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                //DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                //DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                //string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                //string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');


                //sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["cantidad"].ToString().Trim();
                lcantidad += decimal.Parse(row["cantidad"].ToString().Trim());
                //sheet.Cells[lrenglon, lcolumna].numberformat = "0.00";
                sheet.Cells[lrenglon, lcolumna++].value = row["costo"].ToString().Trim(); //Serie Cargo
                lcosto += decimal.Parse(row["costo"].ToString().Trim());

                sheet.Cells[lrenglon, lcolumna++].value = row["TOTAL"].ToString().Trim(); //Fecha Cargo
                ltotal += decimal.Parse(row["TOTAL"].ToString().Trim());
                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C

                sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).Style = "Comma";
                sheet.get_Range("D" + lrenglon.ToString(), "E" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            //sheet.Cells.EntireColumn.AutoFit();
            sheet.get_Range("A" + lrenglon.ToString(), "A" + lrenglon.ToString()).EntireColumn.ColumnWidth = 25;
            sheet.get_Range("B" + lrenglon.ToString(), "B" + lrenglon.ToString()).EntireColumn.ColumnWidth = 55;
            sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;
            sheet.get_Range("D" + lrenglon.ToString(), "D" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;
            sheet.get_Range("E" + lrenglon.ToString(), "E" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;
            
            return;
        }


        public void mReportePedidoFacturaComercial(string mEmpresa, DateTime lfechai, DateTime lfechaf)
>>>>>>> 9a41ea45bd8e9002eb6a577c27983ff67c519b3f
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            //configuracionencabezadoComisiones(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);
            //MyExcel.DisplayAlerts = False 

            lrenglon = configuracionencabezadoComisiones(sheet);
            string lultimocliente = "";
            decimal ltotal1 = 0;
            decimal ltotal2 = 0;

            decimal ltotal015 = 0;
            decimal ltotal1530 = 0;
            decimal ltotal3045 = 0;
            decimal ltotal4560 = 0;
            decimal ltotal60 = 0;
            decimal ltotal015a = 0;
            decimal ltotal1530a = 0;
            decimal ltotal3045a = 0;
            decimal ltotal4560a = 0;
            decimal ltotal60a = 0;

            long lultimofolio = 0;

            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                int lcolumna = 1;


                if (row["CCODIGOCLIENTE"].ToString().Trim() != lultimocliente && lultimocliente != "")
                {
                    sheet.Cells[lrenglon, 10].value = ltotal1;
                    sheet.Cells[lrenglon, 16].value = ltotal2;

                    sheet.Cells[lrenglon, 19].value = ltotal015;
                    sheet.Cells[lrenglon, 20].value = ltotal1530;
                    sheet.Cells[lrenglon, 21].value = ltotal3045;
                    sheet.Cells[lrenglon, 22].value = ltotal4560;
                    sheet.Cells[lrenglon, 23].value = ltotal60;

                    sheet.Cells[lrenglon, 25].value = ltotal015a;
                    sheet.Cells[lrenglon, 26].value = ltotal1530a;
                    sheet.Cells[lrenglon, 27].value = ltotal3045a;
                    sheet.Cells[lrenglon, 28].value = ltotal4560a;
                    sheet.Cells[lrenglon, 29].value = ltotal60a;

                    lrenglon += 3;
                    ltotal1 = 0;
                    ltotal2 = 0;
                    ltotal015 = 0;
                    ltotal1530 = 0;
                    ltotal3045 = 0;
                    ltotal4560 = 0;
                    ltotal60 = 0;

                    ltotal015a = 0;
                    ltotal1530a = 0;
                    ltotal3045a = 0;
                    ltotal4560a = 0;
                    ltotal60a = 0;
                }
                //else
                //  ltotal1 += decimal.Parse(row["CPENDIENTE"].ToString().Trim());

                if (lultimofolio != long.Parse(row["CFOLIO"].ToString().Trim()))
                {
                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOCLIENTE"].ToString().Trim();
                    lultimocliente = row["CCODIGOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CDIASCREDITOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CSERIEDOCUMENTO"].ToString().Trim(); //Fecha Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Fecha Cargo
                    string sFecha1 = "'" + row["cfechad"].ToString().Trim();
                    sFecha1 = sFecha1.Substring(0, 11);
                    sheet.Cells[lrenglon, lcolumna++].value = sFecha1;
                    sFecha1 = "'" + row["cfechap"].ToString().Trim();
                    if (sFecha1 != "'")
                    {
                        sFecha1 = sFecha1.Substring(0, 11);
                        sheet.Cells[lrenglon, lcolumna++].value = sFecha1; //Fecha Cargo
                    }
                    else
                        lcolumna++;
                    sheet.Cells[lrenglon, lcolumna++].value = row["CTOTAL"].ToString().Trim(); //Fecha Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CPENDIENTE"].ToString().Trim(); //Fecha Cargo

                    ltotal1 += decimal.Parse(row["CPENDIENTE"].ToString().Trim());

                    sheet.Cells[lrenglon, lcolumna++].value = row["numdias"].ToString().Trim(); //Fecha Cargo
                    lultimofolio = long.Parse(row["CFOLIO"].ToString().Trim());

                }
                else
                    lcolumna += 11;

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADESCAPTURADAS"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CABREVIATURA"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim(); //Fecha Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["ctotalmov"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["comision"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION"].ToString().Trim(); //Fecha Cargo


                //                0 a 15	15 a 30	30 a45	45 a 60	Mayor a 60 dias


                lcolumna = 19;
                /*if (int.Parse(row["numdias"].ToString().Trim()) <= 15 && int.Parse(row["numdias"].ToString().Trim()) > 0)
                {*/
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal015 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                //}

                lcolumna++;
                /*if (int.Parse(row["numdias"].ToString().Trim()) > 15 && int.Parse(row["numdias"].ToString().Trim()) <= 30)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal1530 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);

                }*/
                lcolumna++;
                /*if (int.Parse(row["numdias"].ToString().Trim()) > 30 && int.Parse(row["numdias"].ToString().Trim()) <= 45)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal3045 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }*/
                lcolumna++;
                /*if (int.Parse(row["numdias"].ToString().Trim()) > 45 && int.Parse(row["numdias"].ToString().Trim()) <= 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal4560 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }*/
                lcolumna++;
                /*if (int.Parse(row["numdias"].ToString().Trim()) > 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal60 += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }*/
                lcolumna++;
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) <= 15 && int.Parse(row["numdias"].ToString().Trim()) > 0)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal015a += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 15 && int.Parse(row["numdias"].ToString().Trim()) <= 30)
                {
                    sheet.Cells[lrenglon, lcolumna].value = double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100);
                    ltotal1530a += decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100);

                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 30 && int.Parse(row["numdias"].ToString().Trim()) <= 45)
                {
                    sheet.Cells[lrenglon, lcolumna].value = (double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100)) * 0.75;
                    ltotal3045a += (decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100)) * (decimal)0.75;
                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 45 && int.Parse(row["numdias"].ToString().Trim()) <= 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = (double.Parse(row["ctotalmov"].ToString().Trim()) * (double.Parse(row["comision"].ToString().Trim()) / 100)) * 0.5;
                    ltotal4560a += (decimal.Parse(row["ctotalmov"].ToString().Trim()) * (decimal.Parse(row["comision"].ToString().Trim()) / 100)) * (decimal)0.5;

                }
                lcolumna++;
                if (int.Parse(row["numdias"].ToString().Trim()) > 60)
                {
                    sheet.Cells[lrenglon, lcolumna].value = 0;
                    ltotal60a += 0;
                }

                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                //sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";



                ltotal2 += decimal.Parse(row["ctotalmov"].ToString().Trim());

                lrenglon++;


            }

            sheet.get_Range("I8", "J" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("P8", "P" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("S8", "AC" + lrenglon.ToString()).Style = "Currency";

        }

        public void mReporteCorteCaja(string usuario, string sfecha)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            //configuracionencabezadoComisiones(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);
            //MyExcel.DisplayAlerts = False 

            lrenglon = configuracionencabezadoCorteCaja(sheet,usuario, sfecha);
            string lultimocliente = "";
            
            decimal ltotal1 = 0;
            decimal ltotal2 = 0;

            decimal ltotal3 = 0;
            decimal ltotal4 = 0;
            decimal ltotal5 = 0;
            decimal ltotal6 = 0;
            decimal ltotal60 = 0;
            decimal ltotal015a = 0;
            decimal ltotal1530a = 0;
            decimal ltotal3045a = 0;
            decimal ltotal4560a = 0;
            decimal ltotal60a = 0;

            long lultimofolio = 0;
            string lultimodocumentode = "";
            int lcolumna = 1;
            lrenglon = 5;
            int lrenglonfinal = 0;
            int lrenglonfinal1 = 0;

            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                

                if (row["CIDDOCUMENTODE"].ToString().Trim() != lultimodocumentode && lultimodocumentode != "")
                {
                    if (row["CIDDOCUMENTODE"].ToString().Trim() != "")
                    {
                        lrenglonfinal1 = lrenglon;
                    }
                    lrenglon=5;
                }

                lultimodocumentode = row["CIDDOCUMENTODE"].ToString().Trim();
                //else
                  //  ltotal1 += decimal.Parse(row["CPENDIENTE"].ToString().Trim());

                /*if (lultimofolio != long.Parse(row["CFOLIO"].ToString().Trim()))
                {*/
                if (row["CIDDOCUMENTODE"].ToString().Trim() == "3")
                    lcolumna = 1;
                else
                    lcolumna = 8;

                if (row["CSERIEDOCUMENTO"].ToString().Trim() != "")
                {

                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CSERIEDOCUMENTO"].ToString().Trim();


                    sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Serie Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row[6].ToString().Trim(); //Fecha Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["PROYECTO"].ToString().Trim(); //Fecha Cargo

                    
                    lrenglon++;
                    lrenglonfinal = lrenglon;
                    
                }
                else
                { 
                    // totales

                }
                    

            }


            if (lrenglonfinal1 > 0 || lrenglonfinal > 0)
            {
            if (lrenglonfinal1 > 0)
                sheet.get_Range("E5", "E" + lrenglonfinal1.ToString()).Style = "Currency";
            if (lrenglonfinal > 0)
            sheet.get_Range("L5", "L" + lrenglonfinal.ToString()).Style = "Currency";

            if (lrenglonfinal1 > lrenglonfinal)
                lrenglonfinal = lrenglonfinal1; 
            lrenglon = lrenglonfinal+1;


            sheet.Cells[lrenglon, 5].value = "EFECTIVO";
            sheet.Cells[lrenglon, 6].value = "CHEQUE";
            sheet.Cells[lrenglon, 7].value = "TRANFERENCIA	";
            sheet.Cells[lrenglon, 8].value = "TARJETA";
            sheet.Cells[lrenglon, 9].value = "CREDITO";

            lrenglon++;
            sheet.Cells[lrenglon++, 4].value = "FACTURAS";
            sheet.Cells[lrenglon++, 4].value = "REMISIONES";
            sheet.Cells[lrenglon++, 4].value = "TOTALES";



            lrenglon = lrenglonfinal+1;

            IEnumerable<DataRow> query =
    from product in DatosReporte.AsEnumerable()
    select product;

            IEnumerable<DataRow> largeProducts =
    query.Where(p => p.Field<string>("CSERIEDOCUMENTO") == null);


            string documentode= "";

            //Console.WriteLine("Product Names:");
            foreach (DataRow p in largeProducts)
            {

                lrenglon = lrenglonfinal + 1;

                string proyecto = p.Field<string>("proyecto").ToUpper();
                try
                {
                    documentode = p.Field<Int32>("ciddocumentode").ToString();
                }
                catch (Exception eee)
                {
                    documentode = "0";
                }

                switch (documentode)
                {
                    case "3":
                        lrenglon += 2; break;
                    case "4":
                        lrenglon += 1; break;
                    case "0":
                        lrenglon += 3; break;
                }

                switch (proyecto)
                {
                    case "EFECTIVO":
                        lcolumna = 5; break;
                    case "CHEQUE":
                        lcolumna = 6; break;
                    case "TRANSFERENCIA":
                        lcolumna = 7; break;
                    case "TARJETA":
                        lcolumna = 8; break;
                    case "CREDITO":
                        lcolumna = 9; break;
                }

                if (proyecto != "CANCELADA")
                {
                    string xx = p.Field<double>("total").ToString();

                    sheet.Cells[lrenglon, lcolumna].value = xx;
                    sheet.Cells[lrenglon, lcolumna].Style = "Currency";
                    //sheet.get_Range("L5", "L" + lrenglonfinal.ToString()).Style = "Currency";
                }
            }

            
            }

            lrenglon = lrenglonfinal + 7;
            sheet.Cells[lrenglon, 5].value = "KG";
            sheet.Cells[lrenglon, 6].value = "LT";
            sheet.Cells[lrenglon, 7].value = "SC";
            sheet.Cells[lrenglon, 8].value = "PZ";


            sheet.Cells[lrenglon + 1, 4].value = "FACTURAS";
            sheet.Cells[lrenglon + 2, 4].value = "REMISIONES";
            sheet.Cells[lrenglon + 3, 4].value = "TOTALES";

            lrenglon++;
            lcolumna = 5;
            int lrengloncito = 0;
            foreach (DataRow row in DatosDetalle.Rows)
            {
                switch (row[0].ToString())
                {
                    case "3":
                        lrengloncito = lrenglon;
                        break;
                    case "4":
                        lrengloncito = lrenglon+1;
                        break;
                    case "0":
                        lrengloncito = lrenglon+2;
                        break;
                }
                switch (row["CABREVIATURA"].ToString())
                {
                    case "KG":
                        lcolumna = 5;
                        break;
                    case "LT":
                        lcolumna = 6;
                        break;
                    case "SC":
                        lcolumna = 7;
                        break;
                    case "PZ":
                        lcolumna = 8;
                        break;
                }

                sheet.Cells[lrengloncito, lcolumna].value = decimal.Parse(row[1].ToString()).ToString();

            }

            /*sheet.get_Range("I8" , "J" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("P8", "P" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("S8", "AC" + lrenglon.ToString()).Style = "Currency";*/

            


            
        }



        public void mReporteCobranza(string usuario, string sfecha)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoCobranza(sheet,sfecha);
            //MyExcel.DisplayAlerts = False 
            // c.CCODIGOCLIENTE, c.CRAZONSOCIAL, c.CDIASCREDITOCLIENTE, d.CSERIEDOCUMENTO, d.CFOLIO,d.cfecha, 0,  a.CNOMBREAGENTE, d.CUSUARIO, d.ctotal, d.CPENDIENTE
            //lrenglon = configuracionencabezadoCobranza(sheet, usuario, sfecha);
            string lultimocliente = "";

            

            long lultimofolio = 0;
            string lultimodocumentode = "";
            int lcolumna = 1;
            lrenglon = 6;
            int lrenglonfinal1 = 0;


            decimal newacumulado = 0;
            decimal lacumulado = 0;
            decimal lacumulado2 = 0;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                

                if (row["CCODIGOCLIENTE"].ToString().Trim() != lultimocliente && lultimocliente != "")
                {
                    // cambio de cliente
                    //sheet.Cells[lrenglon, 11].value = lacumulado.ToString().Trim();
                    //sheet.Cells[lrenglon, 11].value = newacumulado.ToString().Trim();
                    
                    lrenglon -= 1;
                    
                    sheet.Cells[lrenglon, 11].value = newacumulado.ToString().Trim();
                    lacumulado = 0;
                    newacumulado = 0;
                    lrenglon++;
                    lrenglon++;
                }




                if (decimal.Parse(row["saldoinicial"].ToString().Trim()) != 0 && lultimocliente != row["CCODIGOCLIENTE"].ToString().Trim())
                {
                    // saldo inicial
                    lcolumna = 1;
                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CRAZONSOCIAL"].ToString().Trim();
                    lcolumna += 5;
                    sheet.Cells[lrenglon, lcolumna++].value = "SALDO INICIAL";

                    if (row["cnombreagentesaldo"].ToString().Trim() != "")
                        sheet.Cells[lrenglon, 9].value = row["cnombreagentesaldo"].ToString().Trim();


                    lcolumna += 1;
                    sheet.Cells[lrenglon, lcolumna++].value = row["saldoinicial"].ToString().Trim();
                    //sheet.Cells[lrenglon, lcolumna++].value = row["saldoinicial"].ToString().Trim();
                    lrenglon++;
                    lacumulado = decimal.Parse(row["saldoinicial"].ToString().Trim());
                    newacumulado = decimal.Parse(row["saldoinicial"].ToString().Trim());
                }

                lultimocliente = row["CCODIGOCLIENTE"].ToString().Trim();

                lcolumna = 1;
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOCLIENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CRAZONSOCIAL"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CDIASCREDITOCLIENTE"].ToString().Trim();
              //  sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Serie Cargo
                
                sheet.Cells[lrenglon, lcolumna++].value = row["CSERIEDOCUMENTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CFECHA"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["diasdif"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim();

                

                sheet.Cells[lrenglon, lcolumna++].value = row["CUSUARIO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim(); // el pendiente queda como el total

               // sheet.Cells[lrenglon, lcolumna++].value = row["cpendiente"].ToString().Trim();


                   // , d.CFOLIO,d.cfecha, 0,  a.CNOMBREAGENTE, d.CUSUARIO, d.ctotal, d.CPENDIENTE"].ToString().Trim(); //Serie Cargo
                newacumulado += decimal.Parse(row["cpendiente"].ToString().Trim());

                lacumulado = decimal.Parse(row["acumulado"].ToString().Trim()) + decimal.Parse(row["saldoinicial"].ToString().Trim());
                    lrenglon++;


            }
            sheet.get_Range("J6", "K" + lrenglon.ToString()).Style = "Currency";

            

            //sheet.get_Range("I" + lrenglon.ToString(), "N" + lrenglon.ToString()).Style = "Currency";

            




         
    



        }



        public void mReporteRemisionesComercial(string mEmpresa, DateTime lfechai, DateTime lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoRemisionesComercial(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosDetalle.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                //DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                //DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                //string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                //string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');


                //sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADES"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNETO"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CTOTAL"].ToString().Trim(); //Fecha Cargo
                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }

        public void mReportePedidoFacturaComercial(string mEmpresa, DateTime lfechai, DateTime lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoPedidoFacturaComercial(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');


                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOCLIENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOAGENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["foliopedido"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                sheet.Cells[lrenglon, lcolumna++].value = row["foliofactura"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //C

                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION2"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION3"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION4"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION5"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["CVALORCLASIFICACION6"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREMONEDA"].ToString().Trim(); //pendiente de facturar

                sheet.Cells[lrenglon, lcolumna++].value = row["CPRECIOCAPTURADO"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CPORCENTAJEDESCUENTO1"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADESCAPTURADAS"].ToString().Trim(); //pendiente de facturar
                sheet.Cells[lrenglon, lcolumna++].value = row["cneto"].ToString().Trim(); //pendiente de facturar


                sheet.Cells[lrenglon, lcolumna++].value = row["CDESCUENTO1"].ToString().Trim(); //pendiente de facturar
                sheet.Cells[lrenglon, lcolumna++].value = row["CIMPUESTO1"].ToString().Trim(); //pendiente de facturar


                sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fechav;

                sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }

        public void mReporteZonaCostoComercial(string mEmpresa, DateTime lfechai, string lnombrealmacen1, string lnombrealmacen2)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoZonaCostoComercial(sheet, mEmpresa, "Zonas y Costos", lrenglon, lfechai,lnombrealmacen1, lnombrealmacen2);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                
                sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["cantidad"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREALMACEN"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["CZONA"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["costo"].ToString().Trim(); //C
                sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).NumberFormat = "#,##0.00_ ;-#,##0.00 ";

                sheet.get_Range("F" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }

        public void mReporteVentasLotesComercial(string mEmpresa)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoVentasLotesComercial(sheet, mEmpresa, lrenglon);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            int lcolumna=1;
            int lidmovimiento = 0;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //SERIE	FOLIO	CODIGO CLIENTE	RAZON SOCIAL	NOMBRE AGENTE	FECHA	CANCELADO	UNIDADES	CODIGO PRODUCTO	NOMBRE PRODUCTO	PRECIO	Neto-Desc	TOTAL	CONCEPTO DOCUMENTO	REFERENCIA	TEXTO EXTRA1	TOTAL UNIDADES	OBSERVACIONES	NUMERO LOTE
                if (lidmovimiento == int.Parse(row["CIDMOVIMIENTO"].ToString().Trim()))
                {
                    lrenglon--;
                    // llenar otro lote
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNUMEROLOTE"].ToString().Trim(); //C
                    
                }
                /*d.CSERIEDOCUMENTO
, d.CFOLIO
, c.CCODIGOCLIENTE
, c.CRAZONSOCIAL
, a.CNOMBREAGENTE
, d.CFECHA
, d.CCANCELADO
, m.cunidades
, p.CCODIGOPRODUCTO
, p.CNOMBREPRODUCTO
, m.CPRECIO
, m.cneto - (m.CDESCUENTO1 + m.CDESCUENTO2 + m.CDESCUENTO3 + m.CDESCUENTO4 + m.CDESCUENTO5)
,m.ctotal
, co.CCODIGOCONCEPTO
, m.CREFERENCIA
, m.CTEXTOEXTRA1
, m.COBSERVAMOV
, ca.CNUMEROLOTE
, ca.CEXISTENCIA
, m.CIDMOVIMIENTO*/
                else
                {
                    lcolumna = 1;

                    sheet.Cells[lrenglon, lcolumna++].value = row["CSERIEDOCUMENTO"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOCLIENTE"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim(); //Fecha Cargo



                    string lfecha = row["CFECHA"].ToString().Trim(); //C

                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CFECHA"].ToString().Trim(); //C
                    
                    
                    sheet.Cells[lrenglon, lcolumna++].value = row["CCANCELADO"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["cunidades"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CPRECIO"].ToString().Trim(); //Fecha Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["cneto"].ToString().Trim(); //C
                    sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBRECONCEPTO"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CREFERENCIA"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CTEXTOEXTRA1"].ToString().Trim(); //Serie Cargo

                    sheet.Cells[lrenglon, lcolumna++].value = row["CTOTALUNIDADES"].ToString().Trim();

                    sheet.Cells[lrenglon, lcolumna++].value = row["COBSERVACIONES"].ToString().Trim(); //Fecha Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNUMEROLOTE"].ToString().Trim(); //C
                    if (lrenglon%2==0)
                        sheet.get_Range("A" + lrenglon.ToString(), "S" +
                        lrenglon.ToString()).Interior.Color = Color.LightGreen;

                    //sheet.get_Range("A" + lrenglon.ToString(), "S" +
                    //lrenglon.ToString()).Font.Color = Color.White;

                }

                //sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).NumberFormat = "#,##0.00_ ;-#,##0.00 ";

                sheet.get_Range("K" + lrenglon.ToString(), "M" + lrenglon.ToString()).Style = "Currency";
                lidmovimiento = int.Parse(row["CIDMOVIMIENTO"].ToString().Trim());
                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            sheet.Cells.EntireRow.AutoFit();
            return;
        }



        public void mReporteVentasSNComercial(string mEmpresa)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoVentasSNComercial(sheet, mEmpresa, lrenglon);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            int lcolumna = 1;
            int lidmovimiento = 0;


            

                SqlConnection lconexionconta = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            
            string sempresa2 = Properties.Settings.Default.RutaEmpresaC;

            string server2 = Properties.Settings.Default.server2;
            string user2 = Properties.Settings.Default.user2;
            string pwd2 = Properties.Settings.Default.password2;
            
            
            lconexionconta = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server2 + ";initial catalog = " + sempresa2 + ";user id = " + user2 + "; password = " + pwd2 + ";";
            lconexionconta.ConnectionString = Cadenaconexion1;
            lconexionconta.Open();




            DataSet ds = new DataSet();

            string lsql = "SELECT [Id],[Codigo],[Nombre] from [SegmentosNegocio]";
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, lconexionconta);



            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;
            //mySqlDataAdapter.SelectCommand.CommandText = lsql;

            mySqlDataAdapter.Fill(ds);

            DataTable DatosSN = ds.Tables[0];
            lconexionconta.Close();


            foreach (DataRow row in DatosReporte.Rows)
            {
                /*c.CNOMBRECONCEPTO 
, d.cfecha
, d.cfolio 
, cl.CRAZONSOCIAL
, p.CCODIGOPRODUCTO
, p.CNOMBREPRODUCTO
, m.CUNIDADES
, m.cneto
, m.CIMPUESTO1
, mo.CNOMBREMONEDA
, d.CTIPOCAMBIO
, m.CSCMOVTO*/
                    lcolumna = 1;

                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBRECONCEPTO"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["cfecha"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["cfolio"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["CRAZONSOCIAL"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CCODIGOPRODUCTO"].ToString().Trim(); //Fecha Cargo



    
                    sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //C


                    sheet.Cells[lrenglon, lcolumna++].value = row["CUNIDADES"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumna++].value = row["cneto"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CIMPUESTO1"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CTOTAL"].ToString().Trim(); //Serie Cargo    
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREMONEDA"].ToString().Trim(); //Serie Cargo
                    sheet.Cells[lrenglon, lcolumna++].value = row["CTIPOCAMBIO"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["cestado2"].ToString().Trim(); //Fecha Cargo

                if (row["CSCMOVTO"].ToString().Trim() != "")
                {
                    string searchExpression = "Codigo = " + row["CSCMOVTO"].ToString().Trim() ;

                    DataRow[] foundRows = DatosSN.Select(searchExpression);

                    //DataRow dataRow = DatosSN.AsEnumerable().FirstOrDefault(r => r["Codigo"] == row["CSCMOVTO"].ToString().Trim());
                    


                    if (foundRows.Count() > 0)
                    {
                        sheet.Cells[lrenglon, lcolumna++].value = foundRows[0]["Nombre"].ToString().Trim(); //C
                    }
                }
                if (lrenglon % 2 == 0)
                        sheet.get_Range("A" + lrenglon.ToString(), "N" +
                        lrenglon.ToString()).Interior.Color = Color.LightGreen;

                //sheet.get_Range("A" + lrenglon.ToString(), "S" +
                //lrenglon.ToString()).Font.Color = Color.White;
                lrenglon++;
                

                //sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).NumberFormat = "#,##0.00_ ;-#,##0.00 ";

                //sheet.get_Range("F" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";
                

            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }


        public void mReporteUsuarios()
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            //configuracionencabezadoPedidoFacturaComercial(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);

            sheet.Cells[2, 1].value = "Reporte de Usuario";

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna = 1;
            int idusuario=0;
            int idusuarioant = 0;
            string sGrupo = "", sGrupoAnterior = "";
            foreach (DataRow row in DatosDetalle.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                string sNombre = "", sNombreLargo = "";
               // string sGrupo = "";
                idusuario = int.Parse(row["idusuario"].ToString().Trim());
                if (idusuario != idusuarioant)
                {
                    string cadenaempresas = "";
                    foreach (DataRow rowm in DatosMaestro.Rows)
                    {
                        int idusuariom = int.Parse(rowm["idusuario"].ToString().Trim());
                        if (idusuariom == idusuario)
                        {
                           
                             sNombre = rowm["NOMBRE"].ToString().Trim();
                             sNombreLargo = rowm["NOMBRELARGO"].ToString().Trim();
                             sGrupo = rowm["GRUPO"].ToString().Trim();
                             if (rowm["PROCESO"].ToString().Trim() == "E0")
                                 cadenaempresas = "Todas las empresas";
                             else
                             {
                                 string lnoempresa = rowm["PROCESO"].ToString().Trim();
                                 int   noempresa = int.Parse(lnoempresa.Substring(lnoempresa.IndexOf('E', 0)+1));
                                 foreach (DataRow rowe in DatosEmpresas.Rows)
                                 {

                                     if (int.Parse(rowe[0].ToString()) == noempresa)
                                     {
                                         cadenaempresas += rowe[1].ToString() + ",";
                                     }
 
                                 }
                             }
                        }
                    }
                    lrenglon += 2;
                    sheet.Cells[lrenglon++, lcolumna].value = "Datos Generales";
                    sheet.Cells[lrenglon++, lcolumna].value = "Nombre " + sNombre;
                    sheet.Cells[lrenglon++, lcolumna].value = "Nombre Largo " + sNombreLargo;
                    sheet.Cells[lrenglon++, lcolumna].value = "Grupo " + sGrupo;
                    sheet.Cells[lrenglon++, lcolumna].value = "Empresas a las que tiene acceso ";
                    sheet.Cells[lrenglon++, lcolumna].value = cadenaempresas;
                    //lrenglon++;

                    lrenglon++;
                    sheet.Cells[lrenglon++, lcolumna].value = "DERECHOS Y BARRA DE ACCESO ";
                    sheet.Cells[lrenglon, lcolumna].value = "Modulo ";
                    sheet.Cells[lrenglon, lcolumna+1].value = "Derechos de Acceso ";
                    sheet.Cells[lrenglon, lcolumna + 2].value = "Visualizar en Barra de Acceso ";
                    
                    idusuarioant = idusuario;
                }
                sGrupo = row["grupo"].ToString().Trim();
                lrenglon++;
                if (sGrupo != sGrupoAnterior)
                {
                    lrenglon++;
                    sGrupoAnterior = sGrupo;
                    sheet.Cells[lrenglon++, lcolumna].value = row["grupo"].ToString().Trim();
                    sheet.Cells[lrenglon-1, lcolumna].Font.Bold = true;
                    //sheet.get_Range("A" + 9, "AE" + lrenglon).Font.Bold = true;
                    lrenglon++;
                }
                sheet.Cells[lrenglon, lcolumna].value = row["Proceso"].ToString().Trim();

                if ( row["Estado"].ToString().Trim() == "2")
                {
                    sheet.Cells[lrenglon, lcolumna+1].value = "Si";
                    sheet.Cells[lrenglon, lcolumna + 2].value = "Si";
                }
                if (row["Estado"].ToString().Trim() == "1")
                {
                    sheet.Cells[lrenglon, lcolumna + 1].value = "Si";
                    sheet.Cells[lrenglon, lcolumna + 2].value = "No";
                }
                if (row["Estado"].ToString().Trim() == "0")
                {
                    sheet.Cells[lrenglon, lcolumna + 1].value = "No";
                    sheet.Cells[lrenglon, lcolumna + 2].value = "No";
                }
                //lrenglon++;




            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }

        //configuracionencabezadoZonaCostoComercial(sheet, mEmpresa, "Zonas y Costos", lrenglon, lnombrealmacen1, lnombrealmacen2);

        private void configuracionencabezadoZonaCostoComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, string lAlmacen1, string lAlmacen2)
        {
            //EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            //sheet.Cells[1, 6].value = "Fecha Corte";
            


            string fecha2 = lfecha1.Day.ToString().PadLeft(2, '0') + "/" + lfecha1.Month.ToString().PadLeft(2, '0') + "/" + lfecha1.Year.ToString().PadLeft(4, '0');
        
           // sheet.Cells[1, 7].value = "'" + fecha2;
            




            sheet.Cells[lrenglon, lcolumna++].value = "Codigo ";

            sheet.Cells[lrenglon, lcolumna++].value = "Descripcion";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad";
            sheet.Cells[lrenglon, lcolumna++].value = "Almacen";
            sheet.Cells[lrenglon, lcolumna++].value = "Ubicacion";
            sheet.Cells[lrenglon, lcolumna++].value = "Costo";
            
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;

            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Font.Bold = true;
            
        }

        private void configuracionencabezadoVentasSNComercial(MyExcel.Worksheet sheet, string Empresa, int lrenglon)
        {
            int lcolumna = 1;

            //    												


            sheet.Cells[lrenglon, lcolumna++].value = "Tipo de Documento ";

            sheet.Cells[lrenglon, lcolumna++].value = "Fecha de Factura";
            sheet.Cells[lrenglon, lcolumna++].value = "Folio";
            sheet.Cells[lrenglon, lcolumna++].value = "Razón Social";
            sheet.Cells[lrenglon, lcolumna++].value = "Código Servicio";
            sheet.Cells[lrenglon, lcolumna++].value = "Servicio";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad";

            sheet.Cells[lrenglon, lcolumna++].value = "Subtotal";
            sheet.Cells[lrenglon, lcolumna++].value = "IVA";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";
            sheet.Cells[lrenglon, lcolumna++].value = "Moneda";
            sheet.Cells[lrenglon, lcolumna++].value = "Tipo de Cambio";
            sheet.Cells[lrenglon, lcolumna++].value = "Estado";
            sheet.Cells[lrenglon, lcolumna++].value = "Centro de Costos";
           


            /*sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            */
            //sheet.get_Range("A" + lrenglon, "F" + lrenglon).Font.Bold = true;

            sheet.get_Range("A" + lrenglon.ToString(), "N" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "N" +
            lrenglon.ToString()).Font.Color = Color.White;



        }



        private void configuracionencabezadoVentasLotesComercial(MyExcel.Worksheet sheet, string Empresa, int lrenglon)
        {
            int lcolumna = 1;

//                                                

            sheet.Cells[lrenglon, lcolumna++].value = "SERIE ";

            sheet.Cells[lrenglon, lcolumna++].value = "FOLIO";
            sheet.Cells[lrenglon, lcolumna++].value = "CODIGO CLIENTE";
            sheet.Cells[lrenglon, lcolumna++].value = "RAZON SOCIAL";
            sheet.Cells[lrenglon, lcolumna++].value = "NOMBRE AGENTE";
            sheet.Cells[lrenglon, lcolumna++].value = "FECHA";
            sheet.Cells[lrenglon, lcolumna++].value = "CANCELADO";

            sheet.Cells[lrenglon, lcolumna++].value = "UNIDADES";
            sheet.Cells[lrenglon, lcolumna++].value = "CODIGO PRODUCTO";
            sheet.Cells[lrenglon, lcolumna++].value = "NOMBRE PRODUCTO";
            sheet.Cells[lrenglon, lcolumna++].value = "PRECIO";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto - Desc";
            sheet.Cells[lrenglon, lcolumna++].value = "TOTAL";
            sheet.Cells[lrenglon, lcolumna++].value = "CONCEPTO DOCUMENTO";
            sheet.Cells[lrenglon, lcolumna++].value = "REFERENCIA";
            sheet.Cells[lrenglon, lcolumna++].value = "TEXTO EXTRA1";

            sheet.Cells[lrenglon, lcolumna++].value = "TOTAL UNIDADES";

            sheet.Cells[lrenglon, lcolumna++].value = "OBSERVACIONES";
            sheet.Cells[lrenglon, lcolumna++].value = "NUMERO LOTE";


            /*sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "F" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            */
            //sheet.get_Range("A" + lrenglon, "F" + lrenglon).Font.Bold = true;

            sheet.get_Range("A" + lrenglon.ToString(), "S" +
            lrenglon.ToString()).Interior.Color = Color.Green;

            sheet.get_Range("A" + lrenglon.ToString(), "S" +
            lrenglon.ToString()).Font.Color = Color.White;



        }

        private void configuracionencabezadoPedidoFacturaComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
        {
            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";



            string fecha2 = lfecha1.Day.ToString().PadLeft(2, '0') + "/" + lfecha1.Month.ToString().PadLeft(2, '0') + "/" + lfecha1.Year.ToString().PadLeft(4, '0');
            string fecha3 = lfecha2.Day.ToString().PadLeft(2, '0') + "/" + lfecha2.Month.ToString().PadLeft(2, '0') + "/" + lfecha2.Year.ToString().PadLeft(4, '0');

            sheet.Cells[1, 7].value = "'" + fecha2;
            sheet.Cells[2, 7].value = "'" + fecha3;



            //            Codigo cliente	Descr Cliente	Codigo agente	Desc agente	pedido	Fecha factura	Nombre factura	codigo articolo (principal)	desc articolo	
            //Classificacions	moneda ($ orPesos)	precio unitario	quantidad	precio total	TOTAL	fecha vencimiento



            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Cliente / Codice Cliente";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre Cliente / Cliente";
            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Agente / Codice Agente";
            sheet.Cells[lrenglon, lcolumna++].value = "Agente / Agente";
            sheet.Cells[lrenglon, lcolumna++].value = "Pedido / Comanda";
            sheet.Cells[lrenglon, lcolumna++].value = "Fecha Factura / Data Fattura";

            sheet.Cells[lrenglon, lcolumna++].value = "Número Factura / Numero Fattura";

            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Articulo / Codice Articolo";
            sheet.Cells[lrenglon, lcolumna++].value = "Descripcion Articulo / Descrizione Articolo";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 1/ Classificazione 1";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 2/ Classificazione 2";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 3/ Classificazione 3";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 4/ Classificazione 4";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 5/ Classificazione 5";
            sheet.Cells[lrenglon, lcolumna++].value = "Clasificación 6/ Classificazione 6";
            sheet.Cells[lrenglon, lcolumna++].value = "Moneda / Moneta";
            sheet.Cells[lrenglon, lcolumna++].value = "Precio Unitario / Prezzo Unitario";

            sheet.Cells[lrenglon, lcolumna++].value = "% Descuento / % Sconto";
            // % Descuento / % Sconto

            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad / Quantità";
            sheet.Cells[lrenglon, lcolumna++].value = "Importe Total / Totale";

            sheet.Cells[lrenglon, lcolumna++].value = "Importe Descuento / Quantità scontata";
            sheet.Cells[lrenglon, lcolumna++].value = "I.V.A";
            //
            //  
            sheet.Cells[lrenglon, lcolumna++].value = "Total Factura / Totale fattura";
            sheet.Cells[lrenglon, lcolumna++].value = "Fecha Vencimiento / Data di Scadenza";

            sheet.get_Range("A" + lrenglon, "X" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "X" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "X" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "X" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
        }

        private void configuracionencabezadoRemisionesComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
        {
            int lcolumna = 1;
/*            EncabezadoEmpresa(sheet, Empresa, texto);
            int lcolumna = 1;

            sheet.Cells[1, 6].value = "Fecha Inicial";
            sheet.Cells[2, 6].value = "Fecha Final";



            string fecha2 = lfecha1.Day.ToString().PadLeft(2, '0') + "/" + lfecha1.Month.ToString().PadLeft(2, '0') + "/" + lfecha1.Year.ToString().PadLeft(4, '0');
            string fecha3 = lfecha2.Day.ToString().PadLeft(2, '0') + "/" + lfecha2.Month.ToString().PadLeft(2, '0') + "/" + lfecha2.Year.ToString().PadLeft(4, '0');

            sheet.Cells[1, 7].value = "'" + fecha2;
            sheet.Cells[2, 7].value = "'" + fecha3;

            */

            //            Codigo cliente	Descr Cliente	Codigo agente	Desc agente	pedido	Fecha factura	Nombre factura	codigo articolo (principal)	desc articolo	
            //Classificacions	moneda ($ orPesos)	precio unitario	quantidad	precio total	TOTAL	fecha vencimiento



            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre";
            sheet.Cells[lrenglon, lcolumna++].value = "Unidades";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
        }





        private void mEncabezadoCelda(MyExcel.Worksheet sheet, string inicio, string fin ,int lrenglon, int lagregarenglon, int tamano, string texto, Boolean Bold = true )
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
        
        //private void configuracionencabezadoComisioens(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
        private int configuracionencabezadoComisiones(MyExcel.Worksheet sheet)
        {
            int lcolumna = 1;
            int lrenglon = 1;

            sheet.Cells[lrenglon++, 2].value = "FERQUIAGRO SA DE CV";
            sheet.Cells[lrenglon++, 2].value = "Reporte de Ventas Detallado";
            //sheet.Cells[lrenglon, 2].value = "FERQUIAGRO SA DE CV";

            lrenglon++;

            

            /*
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            */

            /*sheet.Cells[lrenglon, 1].value = ;
            sheet.get_Range("A" + lrenglon, "K" + (lrenglon+1) ).Merge();*/

            mEncabezadoCelda(sheet, "A", "K", lrenglon,1,16, "Reporte de Cobranza",false );

            mEncabezadoCelda(sheet, "L", "R", lrenglon, 1, 16, "Reporte de Ventas",false);

            mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "Reporte de Comision OBJETIVA");
            mEncabezadoCelda(sheet, "Y", "AC", lrenglon, 0, 11, "Reporte de Comision Proporcional");


            lrenglon++;
            mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "Informacion de comision Porcentaje");
            mEncabezadoCelda(sheet, "Y", "AC", lrenglon, 0, 11, "Informacion de comision Porcentaje");

            lrenglon++;

            mEncabezadoCelda(sheet, "A", "D", lrenglon, 0, 11, "Datos del Cliente");
            mEncabezadoCelda(sheet, "E", "K", lrenglon, 0, 11, "Datos de la Factura");
            mEncabezadoCelda(sheet, "L", "R", lrenglon, 0, 11, "Informacion del Producto");

            mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "100%");

            mEncabezadoCelda(sheet, "Y", "Y", lrenglon, 0, 11, "100%");
            mEncabezadoCelda(sheet, "Z", "Z", lrenglon, 0, 11, "100%");
            mEncabezadoCelda(sheet, "AA", "AA", lrenglon, 0, 11, "75%");
            mEncabezadoCelda(sheet, "AB", "AB", lrenglon, 0, 11, "50%");
            mEncabezadoCelda(sheet, "AC", "AC", lrenglon, 0, 11, "0%");

            lrenglon++;

            mEncabezadoCelda(sheet, "A", "A", lrenglon, 0, 11, "Codigo Cliente");
            mEncabezadoCelda(sheet, "B", "B", lrenglon, 0, 11, "Nombre del Cliente");


            mEncabezadoCelda(sheet, "C", "C", lrenglon, 0, 11, "Dias  Crédito	");
            mEncabezadoCelda(sheet, "D", "D", lrenglon, 0, 11, "Agente de Ventas	");
            mEncabezadoCelda(sheet, "E", "E", lrenglon, 0, 11, "Serie	 ");
            mEncabezadoCelda(sheet, "F", "F", lrenglon, 0, 11, "Folio");
            mEncabezadoCelda(sheet, "G", "G", lrenglon, 0, 11, "Fecha de Emisión");
            mEncabezadoCelda(sheet, "H", "H", lrenglon, 0, 11, "Fecha de Pago");
            mEncabezadoCelda(sheet, "I", "I", lrenglon, 0, 11, "Total Documento");
            mEncabezadoCelda(sheet, "J", "J", lrenglon, 0, 11, "Saldo Vencido");
            mEncabezadoCelda(sheet, "K", "K", lrenglon, 0, 11, "Dias Transcurridos");
            mEncabezadoCelda(sheet, "L", "L", lrenglon, 0, 11, "CANTIDAD");
            mEncabezadoCelda(sheet, "M", "M", lrenglon, 0, 11, "Unidad Base");
            mEncabezadoCelda(sheet, "N", "N", lrenglon, 0, 11, "Codigo");
            mEncabezadoCelda(sheet, "O", "O", lrenglon, 0, 11, "NOMBRE PRODUCTO");
            mEncabezadoCelda(sheet, "P", "P", lrenglon, 0, 11, "TOTAL");
            mEncabezadoCelda(sheet, "Q", "Q", lrenglon, 0, 11, "Comision");
            mEncabezadoCelda(sheet, "R", "R", lrenglon, 0, 11, "Clasificacion");
            mEncabezadoCelda(sheet, "S", "S", lrenglon, 0, 11, "0 a 15");
            mEncabezadoCelda(sheet, "T", "T", lrenglon, 0, 11, "15 a 30");
            mEncabezadoCelda(sheet, "U", "U", lrenglon, 0, 11, "30 a45");
            mEncabezadoCelda(sheet, "V", "V", lrenglon, 0, 11, "45 a 60");
            mEncabezadoCelda(sheet, "W", "W", lrenglon, 0, 11, "Mayor a 60 dias");

            mEncabezadoCelda(sheet, "Y", "Y", lrenglon, 0, 11, "0 a 15");
            mEncabezadoCelda(sheet, "Z", "Z", lrenglon, 0, 11, "15 a 30");
            mEncabezadoCelda(sheet, "AA", "AA", lrenglon, 0, 11, "30 a45");
            mEncabezadoCelda(sheet, "AB", "AB", lrenglon, 0, 11, "45 a 60");
            mEncabezadoCelda(sheet, "AC", "AC", lrenglon, 0, 11, "Mayor a 60 dias");

            lrenglon++;

            return lrenglon;





            /*
            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre";
            sheet.Cells[lrenglon, lcolumna++].value = "Unidades";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
             * */
        }

        private int configuracionencabezadoCobranza(MyExcel.Worksheet sheet, string sFecha)
        {
            int lcolumna = 1;
            int lrenglon = 1;

          /*  sheet.Cells[lrenglon++, 2].value = "FERQUIAGRO SA DE CV";
            sheet.Cells[lrenglon++, 2].value = "Reporte de Ventas Detallado";*/
            //sheet.Cells[lrenglon, 2].value = "FERQUIAGRO SA DE CV";

            lrenglon++;



            /*
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            */

            /*sheet.Cells[lrenglon, 1].value = ;
            sheet.get_Range("A" + lrenglon, "K" + (lrenglon+1) ).Merge();*/
            lrenglon = 2;
            mEncabezadoCelda(sheet, "D", "H", lrenglon, 0, 16, "Reporte de Cobranza", false);
            lrenglon++;
            mEncabezadoCelda(sheet, "E", "G", lrenglon, 0, 12, "del 01/01/" + sFecha + " al " + DateTime.Today.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Today.Month.ToString().PadLeft(2, '0') + "/" + sFecha, false);
            lrenglon++;

            mEncabezadoCelda(sheet, "A", "C", lrenglon, 0, 11, "Informacion Cliente");

            
            mEncabezadoCelda(sheet, "D", "K", lrenglon, 0, 11, "Informacion del Cliente");


            lrenglon++;


            mEncabezadoCelda(sheet, "A", "A", lrenglon, 0, 11, "Codigo Cliente");
            mEncabezadoCelda(sheet, "B", "B", lrenglon, 0, 11, "Nombre del Cliente");


            mEncabezadoCelda(sheet, "C", "C", lrenglon, 0, 11, "Dias  Crédito	");
            mEncabezadoCelda(sheet, "D", "D", lrenglon, 0, 11, "Serie	 ");
            mEncabezadoCelda(sheet, "E", "E", lrenglon, 0, 11, "Folio");
            mEncabezadoCelda(sheet, "F", "F", lrenglon, 0, 11, "Fecha de Emisión");
            mEncabezadoCelda(sheet, "G", "G", lrenglon, 0, 11, "Dias transcurridos");
            mEncabezadoCelda(sheet, "H", "H", lrenglon, 0, 11, "Agente de Ventas Documento	");
            mEncabezadoCelda(sheet, "I", "I", lrenglon, 0, 11, "Agente Sistema	");
            mEncabezadoCelda(sheet, "J", "J", lrenglon, 0, 11, "Total Documento");
            mEncabezadoCelda(sheet, "K", "K", lrenglon, 0, 11, "Saldo Vencido");

            lrenglon++;

            return lrenglon;





            /*
            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre";
            sheet.Cells[lrenglon, lcolumna++].value = "Unidades";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
             * */
        }

<<<<<<< HEAD
        private int configuracionencabezadoCorteCaja(MyExcel.Worksheet sheet, string usuario, string fecha)
=======


        private void configuracionencabezadoForrajeraComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
        {
            int lcolumna = 1;
           //EncabezadoEmpresa(sheet, Empresa, texto);

            sheet.Cells[1,5].value = texto;
            string fecha2 = lfecha1.Day.ToString().PadLeft(2, '0') + "/" + lfecha1.Month.ToString().PadLeft(2, '0') + "/" + lfecha1.Year.ToString().PadLeft(4, '0');
            string fecha3 = lfecha2.Day.ToString().PadLeft(2, '0') + "/" + lfecha2.Month.ToString().PadLeft(2, '0') + "/" + lfecha2.Year.ToString().PadLeft(4, '0');

            

            sheet.Cells[2, 5].value = "Fecha del " + fecha2 + " al " + fecha3;


            lrenglon += 4;
            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Font.Color = Color.White;

            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad";
            sheet.Cells[lrenglon, lcolumna++].value = "Costo Unitario";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
        }
        private void configuracionencabezadoRemisionesComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
>>>>>>> 9a41ea45bd8e9002eb6a577c27983ff67c519b3f
        {
            int lcolumna = 1;
            int lrenglon = 1;

         /*   sheet.Cells[lrenglon++, 2].value = "FERQUIAGRO SA DE CV";
            sheet.Cells[lrenglon++, 2].value = "Reporte de Ventas Detallado";
            //sheet.Cells[lrenglon, 2].value = "FERQUIAGRO SA DE CV";

            lrenglon++;*/



            /*
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            */

            /*sheet.Cells[lrenglon, 1].value = ;
            sheet.get_Range("A" + lrenglon, "K" + (lrenglon+1) ).Merge();*/

            mEncabezadoCelda(sheet, "E", "K", lrenglon, 0, 12, "Corte de Caja", false);

            
            mEncabezadoCelda(sheet, "A", "D", lrenglon, 0, 10, fecha, false);
            lrenglon++;
            mEncabezadoCelda(sheet, "A", "M", lrenglon, 0, 10, "Usuario: " + usuario, false);
            lrenglon++;
            mEncabezadoCelda(sheet, "A", "F", lrenglon, 0, 10, "Remisiones ", false);
            mEncabezadoCelda(sheet, "H", "M", lrenglon,0, 10, "Facturas ", false);

            lrenglon++;


            mEncabezadoCelda(sheet, "A", "A", lrenglon, 0, 10, "SERIE", false);
            mEncabezadoCelda(sheet, "B", "B", lrenglon, 0, 10, "FOLIO", false);
            mEncabezadoCelda(sheet, "C", "C", lrenglon, 0, 10, "CLAVE", false);
            mEncabezadoCelda(sheet, "D", "D", lrenglon, 0, 10, "NOMBRE", false);
            mEncabezadoCelda(sheet, "E", "E", lrenglon, 0, 10, "MONTO", false);
            mEncabezadoCelda(sheet, "F", "F", lrenglon, 0, 10, "PROYECTO", false);


            mEncabezadoCelda(sheet, "H", "H", lrenglon, 0, 10, "SERIE", false);
            mEncabezadoCelda(sheet, "I", "I", lrenglon, 0, 10, "FOLIO", false);
            mEncabezadoCelda(sheet, "J", "J", lrenglon, 0, 10, "CLAVE", false);
            mEncabezadoCelda(sheet, "K", "K", lrenglon, 0, 10, "NOMBRE", false);
            mEncabezadoCelda(sheet, "L", "L", lrenglon, 0, 10, "MONTO", false);
            mEncabezadoCelda(sheet, "M", "M", lrenglon, 0, 10, "PROYECTO", false);

            lrenglon++;

            return lrenglon;
        }


        private int configuracionencabezadoComisionesRemision(MyExcel.Worksheet sheet)
        {
            int lcolumna = 1;
            int lrenglon = 1;

            sheet.Cells[lrenglon++, 2].value = "FERQUIAGRO SA DE CV";
            sheet.Cells[lrenglon++, 2].value = "Reporte de Ventas Detallado";
            //sheet.Cells[lrenglon, 2].value = "FERQUIAGRO SA DE CV";

            lrenglon++;



            /*
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "A" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            */

            /*sheet.Cells[lrenglon, 1].value = ;
            sheet.get_Range("A" + lrenglon, "K" + (lrenglon+1) ).Merge();*/

            //mEncabezadoCelda(sheet, "A", "K", lrenglon, 1, 16, "Reporte de Cobranza", false);

            mEncabezadoCelda(sheet, "A", "O", lrenglon, 1, 16, "Reporte de Ventas", false);

            /*mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "Reporte de Comision OBJETIVA");
            mEncabezadoCelda(sheet, "Y", "AC", lrenglon, 0, 11, "Reporte de Comision Proporcional");


            lrenglon++;
            mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "Informacion de comision Porcentaje");
            mEncabezadoCelda(sheet, "Y", "AC", lrenglon, 0, 11, "Informacion de comision Porcentaje");
            */
            lrenglon++;

            mEncabezadoCelda(sheet, "A", "C", lrenglon, 0, 11, "Datos del Cliente");
            mEncabezadoCelda(sheet, "D", "G", lrenglon, 0, 11, "Datos de la Remision");
            mEncabezadoCelda(sheet, "H", "O", lrenglon, 0, 11, "Informacion del Producto");

           /* mEncabezadoCelda(sheet, "S", "W", lrenglon, 0, 11, "100%");

            mEncabezadoCelda(sheet, "Y", "Y", lrenglon, 0, 11, "100%");
            mEncabezadoCelda(sheet, "Z", "Z", lrenglon, 0, 11, "100%");
            mEncabezadoCelda(sheet, "AA", "AA", lrenglon, 0, 11, "75%");
            mEncabezadoCelda(sheet, "AB", "AB", lrenglon, 0, 11, "50%");
            mEncabezadoCelda(sheet, "AC", "AC", lrenglon, 0, 11, "0%");*/

            lrenglon++;

            mEncabezadoCelda(sheet, "A", "A", lrenglon, 0, 11, "Codigo Cliente");
            mEncabezadoCelda(sheet, "B", "B", lrenglon, 0, 11, "Nombre del Cliente");


            mEncabezadoCelda(sheet, "C", "C", lrenglon, 0, 11, "Agente de Ventas");
            mEncabezadoCelda(sheet, "D", "D", lrenglon, 0, 11, "Serie Remision");
            //mEncabezadoCelda(sheet, "E", "E", lrenglon, 0, 11, "Serie	 ");
            mEncabezadoCelda(sheet, "E", "E", lrenglon, 0, 11, "Folio Remision");
            mEncabezadoCelda(sheet, "F", "F", lrenglon, 0, 11, "Fecha de Emisión");

            mEncabezadoCelda(sheet, "G", "G", lrenglon, 0, 11, "Total Documento");
            //mEncabezadoCelda(sheet, "I", "I", lrenglon, 0, 11, "Total Documento");
            mEncabezadoCelda(sheet, "H", "H", lrenglon, 0, 11, "CANTIDAD");
            //mEncabezadoCelda(sheet, "K", "K", lrenglon, 0, 11, "Dias Transcurridos");
            //mEncabezadoCelda(sheet, "L", "L", lrenglon, 0, 11, "CANTIDAD");
            mEncabezadoCelda(sheet, "I", "I", lrenglon, 0, 11, "Unidad Base");
            mEncabezadoCelda(sheet, "J", "J", lrenglon, 0, 11, "Codigo");
            mEncabezadoCelda(sheet, "K", "K", lrenglon, 0, 11, "NOMBRE PRODUCTO");
            mEncabezadoCelda(sheet, "L", "L", lrenglon, 0, 11, "TOTAL");
            mEncabezadoCelda(sheet, "M", "M", lrenglon, 0, 11, "%Comision");
            mEncabezadoCelda(sheet, "N", "N", lrenglon, 0, 11, "Clasificacion");
            mEncabezadoCelda(sheet, "O", "O", lrenglon, 0, 11, "Monto Comsion");
            /*mEncabezadoCelda(sheet, "T", "T", lrenglon, 0, 11, "15 a 30");
            mEncabezadoCelda(sheet, "U", "U", lrenglon, 0, 11, "30 a45");
            mEncabezadoCelda(sheet, "V", "V", lrenglon, 0, 11, "45 a 60");
            mEncabezadoCelda(sheet, "W", "W", lrenglon, 0, 11, "Mayor a 60 dias");

            mEncabezadoCelda(sheet, "Y", "Y", lrenglon, 0, 11, "0 a 15");
            mEncabezadoCelda(sheet, "Z", "Z", lrenglon, 0, 11, "15 a 30");
            mEncabezadoCelda(sheet, "AA", "AA", lrenglon, 0, 11, "30 a45");
            mEncabezadoCelda(sheet, "AB", "AB", lrenglon, 0, 11, "45 a 60");
            mEncabezadoCelda(sheet, "AC", "AC", lrenglon, 0, 11, "Mayor a 60 dias");
            */
            lrenglon++;

            return lrenglon;

            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Font.Color = Color.White;




            /*
            sheet.Cells[lrenglon, lcolumna++].value = "Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre";
            sheet.Cells[lrenglon, lcolumna++].value = "Unidades";
            sheet.Cells[lrenglon, lcolumna++].value = "Neto";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
             * */
        }
        public void mReporteInventarioCapas(string mEmpresa, string lfechai, string lfechaf)
        {

            //mConfigurarObjetosImpresion();
            mConfigurarObjetosSQLImpresion();


            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            //int lcolumna = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadocapas(sheet, mEmpresa, "", lrenglon, lfechai, lfechaf);


         

            //into a
            //from b in a.DefaultIfEmpty(new Order())

          var all = from lprods in listaprods
                    join lcapasinicial in listacapas on lprods.IdProducto equals lcapasinicial.IdProducto into a
                    from b in a.DefaultIfEmpty(new RegCapas())

          select new 
          {
              nombreproducto1 = lprods.NombreProducto,
              codigo = lprods.CodigoProducto,
              existenciaproducto = lprods.Existencia,
              entradas = lprods.EntradasPeriodo,
            salidas = lprods.SalidasPeriodo,
              existenciainicialcapa = b.ExistenciaInicial,
              existenciaentradacapa = b.ExistenciaEntradasPeriodo,
              existenciasalidacapa = b.ExistenciaSalidasPeriodo,
              //existenciafinalcapa = b.ExistenciaInicial + b.ExistenciaEntradasPeriodo - b.ExistenciaSalidasPeriodo,
              fechacapa = b.Fecha,
              metodo = lprods.metodocosteo,
              costo = b.costo,
              Clasif1 = lprods.Clasif1,
              Clasif2 = lprods.Clasif2,
              Clasif3 = lprods.Clasif3,
              Clasif4 = lprods.Clasif4,
              Clasif5 = lprods.Clasif5,
              Clasif6 = lprods.Clasif6,
              almacen = b.almacen,
              idcapa = b.idcapas
          };

            string lnombre = "";
            //lrenglon=0;
            int lrenglonnormal = lrenglon;
            decimal inventarioinicialimportes = 0;
            decimal inventarioentradasimportes = 0;
            decimal inventariosalidasimportes = 0;
            decimal inventariofinalimportes = 0;
            foreach (var todo in all)
            {
                if (lnombre != todo.nombreproducto1){
                    if (lrenglon != lrenglonnormal)
                    {
                        sheet.Cells[lrenglonnormal, 9].value = inventarioinicialimportes;
                        sheet.Cells[lrenglonnormal, 10].value = inventarioentradasimportes;
                        sheet.Cells[lrenglonnormal, 11].value = inventariosalidasimportes;
                        sheet.Cells[lrenglonnormal, 12].value = inventariofinalimportes;
                        sheet.get_Range("E" + lrenglonnormal.ToString(), "M" + lrenglonnormal.ToString()).Style = "Currency";
                    }
                    inventarioinicialimportes = 0;
                    inventarioentradasimportes = 0;
                    inventariosalidasimportes = 0;
                    inventariofinalimportes = 0;
                    lrenglon++;
                    lrenglonnormal = lrenglon;
                    sheet.Cells[lrenglon, 1].value = "Normal";
                    sheet.Cells[lrenglon, 2].value = todo.codigo;
                    sheet.Cells[lrenglon, 3].value = todo.nombreproducto1;
                    switch (todo.metodo)
                    {
                        case 1: sheet.Cells[lrenglon, 4].value = "Costo Promedio"; break;
                        case 2: sheet.Cells[lrenglon, 4].value = "Costo Promedio por Almacen"; break;
                        case 3: sheet.Cells[lrenglon, 4].value = "Ultimo Costo"; break;
                        case 4: sheet.Cells[lrenglon, 4].value = "UEPS"; break;
                        case 5: sheet.Cells[lrenglon, 4].value = "PEPS"; break;
                        case 6: sheet.Cells[lrenglon, 4].value = "Costo especifico"; break;
                        case 7: sheet.Cells[lrenglon, 4].value = "Costo Estandar"; break;


                    }
                    sheet.Cells[lrenglon, 5].value = todo.existenciaproducto;
                    sheet.Cells[lrenglon, 6].value = todo.entradas;
                    sheet.Cells[lrenglon, 7].value = todo.salidas;
                    sheet.Cells[lrenglon, 8].value = todo.existenciaproducto + todo.entradas - todo.salidas;

                    sheet.Cells[lrenglon, 15].value = todo.Clasif1;
                    sheet.Cells[lrenglon, 16].value = todo.Clasif2;
                    sheet.Cells[lrenglon, 17].value = todo.Clasif3;
                    sheet.Cells[lrenglon, 18].value = todo.Clasif4;
                    sheet.Cells[lrenglon, 19].value = todo.Clasif5;
                    sheet.Cells[lrenglon, 20].value = todo.Clasif6;

                    sheet.Cells[lrenglon, 8].value = todo.existenciaproducto + todo.entradas - todo.salidas;
                    lnombre = todo.nombreproducto1;
                    sheet.get_Range("A" + lrenglon.ToString(), "T" + lrenglon.ToString()).Interior.Color = Color.LightBlue;
                }
                if (todo.metodo == 5){
                    lrenglon++;
                    sheet.Cells[lrenglon, 1].value = "Capas";
                    sheet.Cells[lrenglon, 2].value = todo.codigo;
                    sheet.Cells[lrenglon, 3].value = todo.nombreproducto1;
                    switch (todo.metodo)
                    {
                        case 1: sheet.Cells[lrenglon, 4].value = "Costo Promedio"; break;
                        case 2: sheet.Cells[lrenglon, 4].value = "Costo Promedio por Almacen"; break;
                        case 3: sheet.Cells[lrenglon, 4].value = "Ultimo Costo"; break;
                        case 4: sheet.Cells[lrenglon, 4].value = "UEPS"; break;
                        case 5: sheet.Cells[lrenglon, 4].value = "PEPS"; break;
                        case 6: sheet.Cells[lrenglon, 4].value = "Costo especifico"; break;
                        case 7: sheet.Cells[lrenglon, 4].value = "Costo Estandar"; break;
                        
                    
                }
                    //sheet.Cells[lrenglon, 4].value = todo.fechacapa;
                    sheet.Cells[lrenglon, 5].value = todo.existenciainicialcapa;
                    decimal entradacapa = todo.existenciaentradacapa;
                    sheet.Cells[lrenglon, 6].value = entradacapa;
                    if (todo.existenciainicialcapa > 0){
                        sheet.Cells[lrenglon, 6].value = 0;
                        entradacapa = 0;
                    }
                    sheet.Cells[lrenglon, 7].value = todo.existenciasalidacapa;
                    sheet.Cells[lrenglon, 8].value = todo.existenciainicialcapa + entradacapa - todo.existenciasalidacapa;
                    sheet.Cells[lrenglon, 9].value = todo.existenciainicialcapa * todo.costo;
                    inventarioinicialimportes += (todo.existenciainicialcapa * todo.costo);
                    
                    sheet.Cells[lrenglon, 10].value = entradacapa * todo.costo;
                    inventarioentradasimportes += (entradacapa * todo.costo);

                    sheet.Cells[lrenglon, 11].value = todo.existenciasalidacapa * todo.costo;
                    inventariosalidasimportes += (todo.existenciasalidacapa * todo.costo);
                    sheet.Cells[lrenglon, 12].value = (todo.existenciainicialcapa + entradacapa - todo.existenciasalidacapa) * todo.costo;
                inventariofinalimportes += ((todo.existenciainicialcapa + entradacapa - todo.existenciasalidacapa) * todo.costo);
                    sheet.Cells[lrenglon, 13].value = todo.costo;
                    sheet.Cells[lrenglon, 14].value = todo.almacen;

                    sheet.Cells[lrenglon, 15].value = todo.Clasif1;
                    sheet.Cells[lrenglon, 16].value = todo.Clasif2;
                    sheet.Cells[lrenglon, 17].value = todo.Clasif3;
                    sheet.Cells[lrenglon, 18].value = todo.Clasif4;
                    sheet.Cells[lrenglon, 19].value = todo.Clasif5;
                    sheet.Cells[lrenglon, 20].value = todo.Clasif6;

                    //sheet.Cells[lrenglon, 21].value = todo.idcapa;


                    sheet.get_Range("A" + lrenglon.ToString(), "M" + lrenglon.ToString()).Interior.Color = Color.LightYellow;
                }
                    sheet.get_Range("E" + lrenglon.ToString(), "M" + lrenglon.ToString()).Style = "Currency";
                    

           }

            sheet.Cells[lrenglonnormal, 9].value = inventarioinicialimportes;
            sheet.Cells[lrenglonnormal, 10].value = inventarioentradasimportes;
            sheet.Cells[lrenglonnormal, 11].value = inventariosalidasimportes;
            sheet.Cells[lrenglonnormal, 12].value = inventariofinalimportes;
            sheet.get_Range("E" + lrenglonnormal.ToString(), "M" + lrenglonnormal.ToString()).Style = "Currency";
            sheet.Cells.EntireColumn.AutoFit();
            return;

            
        }


        private void subconfiguracionencabezadofotos(MyExcel.Worksheet sheet, int lrenglon)
        {

            //            Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total

            //FOLIO		FECHA DE PEDIDO		FECHA DE RECEPCIÓN		FECHA DE VENCIMIENTO		MODELO PROVEEDOR		MOD. TIENDA		FOTOGRAFÍA		CANTIDAD		UNIDAD		PRECIO PROVEEDOR		TOTAL		
            //ORIGEN		CLIENTE		PEDIDO DEL CLIENTE		PENDIENTES POR SURTIR

            int lcolumna = 1;
            //lrenglon = 8;

            sheet.get_Range("A" + lrenglon.ToString(), "AE" + lrenglon.ToString()).Interior.Color = Color.DarkBlue;
            sheet.get_Range("A" + lrenglon.ToString(), "AE" + lrenglon.ToString()).Font.Color = Color.White;

            sheet.get_Range("A" + lrenglon.ToString(), "AE" + lrenglon.ToString()).Font.Size = 10;
            

            sheet.Cells[lrenglon, lcolumna++].value = "FOLIO";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "FECHA DE PEDIDO";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "FECHA DE RECEPCIÓN";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "FECHA DE VENCIMIENTO";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "MODELO PROVEEDOR";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "MOD. TIENDA";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "NOMBRE";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "FOTOGRAFÍA";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "CANTIDAD";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "UNIDAD";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "PRECIO PROVEEDOR";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "TOTAL";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "ORIGEN";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "CLIENTE";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "PEDIDO DEL CLIENTE";
            lcolumna++;
            sheet.Cells[lrenglon, lcolumna++].value = "PENDIENTES POR SURTIR";
            lcolumna++;



        
        }


        public void mReporteFotos(string mEmpresa, string rutaarchivoexcel)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel(@rutaarchivoexcel);
            int lrenglon = 3;
            int lrengloninicial = 3;
            int lrengloniniciaconcepto = 3;
            int lrenglontempo = 3;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            

            configuracionencabezadofotos(sheet, mEmpresa, "", lrenglon, "", "");

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon = 8;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna=1;
            string lProveedorPrevio = "";
            DateTime dfecha;
            string lfecha;

            sheet.Range["A9:AH32000"].Delete();

            foreach (DataRow row in DatosReporteAdminpaq.Rows)
            {
                //select m2.ccodigoc01, m2.crazonso01, m8.creferen01,  ;
//m8.cfecha, m8.cfechaen01, m8.cfechave01, m5.cnomaltern, m5.ccodigop01, m10.cunidades, m26.cdesplie01, m10.cprecioc01, m10.ctotal,m10.creferen01, m10.ctextoex02, m10.ctextoex03, m10.cunidade03 ;



                string lProveedor = row["ccodigoc01"].ToString().Trim();
                if (lProveedorPrevio != lProveedor)
                {
                    lcolumna = 1;
                    lrenglon++;

                    sheet.get_Range("A" + lrenglon.ToString(), "G" + lrenglon.ToString()).Interior.Color = Color.DarkBlue;
                    sheet.get_Range("A" + lrenglon.ToString(), "G" + lrenglon.ToString()).Font.Color = Color.White;
                    sheet.Cells[lrenglon, lcolumna++].value = "Proveedor";
                    sheet.Cells[lrenglon++, lcolumna].value = row["ccodigoc01"].ToString().Trim() + "-" + row["crazonso01"].ToString().Trim(); //Folio Cargo
                    lProveedorPrevio = lProveedor;
                    lrenglon++;

                    subconfiguracionencabezadofotos(sheet, lrenglon);
                    lrenglon++;
                }

                lcolumna = 1;


               // sheet.Cells[lrenglon, lcolumna++].value = row["creferen01"].ToString().Trim(); //Folio Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["cseriedo01"].ToString().Trim() + "-" + row["cfolio"].ToString().Trim(); //Folio Cargo

                lcolumna++;

                dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());


                string ssfecha = ""; 
                ssfecha = "'"  + dfecha.ToString("dd-MMM-yy").Replace(".", ""); 

                lfecha = "'" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                lfecha = "'" + dfecha.ToString("d") + "/" + dfecha.ToString("MMM") + "/" + dfecha.ToString("yy");


                sheet.Cells[lrenglon, lcolumna++].value = ssfecha; //Serie Cargo
                lcolumna++;
                dfecha = DateTime.Parse(row["cfechaen01"].ToString().Trim());

                ssfecha = "'" +  dfecha.ToString("dd-MMM-yy").Replace(".", ""); 


                lfecha = "'" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                lfecha = "'" + dfecha.ToString("d") + "/" + dfecha.ToString("MMM") + "/" + dfecha.ToString("yy");




                
                sheet.Cells[lrenglon, lcolumna++].value = ssfecha;
                
                

                lcolumna++;
                dfecha = DateTime.Parse(row["cfechave01"].ToString().Trim());
                ssfecha = "'" +  dfecha.ToString("dd-MMM-yy").Replace(".", ""); 
                lfecha = "'" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                lfecha = "'" + dfecha.ToString("d") + "/" + dfecha.ToString("MMM") + "/" + dfecha.ToString("yy");


                sheet.Cells[lrenglon, lcolumna++].value = ssfecha; //C
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["cnomaltern"].ToString().Trim(); //importe
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["ccodigop01"].ToString().Trim(); //pendiente de facturar
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["cnombrep01"].ToString().Trim(); //pendiente de facturar
                lcolumna++;
                lcolumna++;
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["cunidades"].ToString().Trim(); //serie factura
                lcolumna++;
                
                sheet.Cells[lrenglon, lcolumna++].value = row["cdesplie01"].ToString().Trim(); //# de factura
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["cprecioc01"].ToString().Trim(); //# de factura
                lcolumna++;

                sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim(); //Folio Cargo
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["creferen02"].ToString().Trim(); //Serie Cargo
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["ctextoex02"].ToString().Trim(); //Fecha Cargo
                lcolumna++;
                sheet.Cells[lrenglon, lcolumna++].value = row["ctextoex03"].ToString().Trim(); //importe
                lcolumna++;

                sheet.Cells[lrenglon, lcolumna++].value = row["cunidade03"].ToString().Trim(); //importe
                lcolumna++;

                sheet.get_Range("A" + lrenglon, "AE" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "AE" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "AE" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;

                sheet.Rows[lrenglon].RowHeight =100;

                

                //sheet.get_Range("E" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";
                //sheet.get_Range("J" + lrenglon.ToString(), "M" + lrenglon.ToString()).Style = "Currency";
                lrenglon++;


            }

            sheet.get_Range("A" + 9, "AE" + lrenglon).Font.Bold = true;
            sheet.get_Range("A" + 9, "AE" + lrenglon).Font.Size = 10;

            sheet.Columns["A:A"].ColumnWidth = 15;
            sheet.Columns["B:B"].ColumnWidth = 5;
            sheet.Columns["C:C"].ColumnWidth = 15;
            sheet.Columns["D:D"].ColumnWidth = 1;

            sheet.Columns["E:E"].ColumnWidth = 15;
            sheet.Columns["F:F"].ColumnWidth = 1;

            sheet.Columns["G:G"].ColumnWidth = 20;
            sheet.Columns["H:H"].ColumnWidth = 1;

            sheet.Columns["I:I"].ColumnWidth = 15;
            sheet.Columns["J:J"].ColumnWidth = 1;

            sheet.Columns["K:K"].ColumnWidth = 15;
            sheet.Columns["L:L"].ColumnWidth = 1;


            sheet.Columns["M:M"].ColumnWidth = 57;
            sheet.Columns["N:N"].ColumnWidth = 1;

            sheet.Columns["O:O"].ColumnWidth = 50;
            sheet.Columns["P:P"].ColumnWidth = 1;

            sheet.Columns["O:O"].ColumnWidth = 50;
            sheet.Columns["P:P"].ColumnWidth = 1;

            sheet.Columns["Q:Q"].ColumnWidth = 15;
            sheet.Columns["R:R"].ColumnWidth = 1;

            sheet.Columns["S:S"].ColumnWidth = 15;
            sheet.Columns["T:T"].ColumnWidth = 1;

            sheet.Columns["U:U"].ColumnWidth = 15;
            sheet.Columns["V:V"].ColumnWidth = 1;

            sheet.Columns["W:W"].ColumnWidth = 15;
            sheet.Columns["X:X"].ColumnWidth = 1;

            sheet.Columns["Y:Y"].ColumnWidth = 15;
            sheet.Columns["Z:Z"].ColumnWidth = 1;

            sheet.Columns["AA:AA"].ColumnWidth = 15;
            sheet.Columns["AB:AB"].ColumnWidth = 1;

            sheet.Columns["AC:AC"].ColumnWidth = 15;
            sheet.Columns["AD:AD"].ColumnWidth = 1;

            sheet.Columns["AE:AE"].ColumnWidth = 15;




            //sheet.Cells.EntireColumn.AutoFit();
            return;
        }


        public void mReporteIEPS(string mEmpresa, string lfechai, string lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoieps(sheet, mEmpresa, "", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosFacturaAbono.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total

                lcolumna = 1;
                sheet.Cells[lrenglon, lcolumna++].value = row["Foliocargo"].ToString().Trim(); //Folio Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Seriecargo"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Fechacargo"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Cliente"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["Conceptocargo"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["Netocargo"].ToString().Trim(); //pendiente de facturar
                sheet.Cells[lrenglon, lcolumna++].value = row["Iva"].ToString().Trim(); //serie factura
                sheet.Cells[lrenglon, lcolumna++].value = row["Ieps"].ToString().Trim(); //# de factura
                sheet.Cells[lrenglon, lcolumna++].value = row["Totalcargo"].ToString().Trim(); //# de factura

                sheet.Cells[lrenglon, lcolumna++].value = row["Folioabono"].ToString().Trim(); //Folio Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Serieabono"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Fechaabono"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["Conceptoabono"].ToString().Trim(); //importe

                sheet.Cells[lrenglon, lcolumna++].value = row["referencia"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["observaciones"].ToString().Trim(); //importe


                sheet.Cells[lrenglon, lcolumna++].value = row["Totalabono"].ToString().Trim(); //# de factura
                sheet.Cells[lrenglon, lcolumna++].value = row["pagado"].ToString().Trim(); //# de factura

                decimal totalcargo = decimal.Parse(row["Totalcargo"].ToString().Trim());
                decimal pagado = decimal.Parse(row["pagado"].ToString().Trim());
                decimal netocargo = decimal.Parse(row["Netocargo"].ToString().Trim());
                decimal IVA = decimal.Parse(row["Iva"].ToString().Trim());
                decimal IEPS = decimal.Parse(row["Ieps"].ToString().Trim());
                decimal celda = pagado*netocargo/totalcargo;




                sheet.Cells[lrenglon, lcolumna++].value = celda.ToString().Trim(); //# de factura

                celda = pagado * IVA / totalcargo;
                sheet.Cells[lrenglon, lcolumna++].value = celda.ToString().Trim(); //# de factura

                celda = pagado * IEPS / totalcargo;
                sheet.Cells[lrenglon, lcolumna++].value = celda.ToString().Trim(); //# de factura
                
                



                //sheet.get_Range("E" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";
                //sheet.get_Range("J" + lrenglon.ToString(), "M" + lrenglon.ToString()).Style = "Currency";
                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }



        public void mReporteOC(string mEmpresa, string lfechai, string lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadooc(sheet, mEmpresa, "Productos Pendientes de Surtir Ordenes de Compras", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosFacturaAbono.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"
 
                lcolumna = 1;
                sheet.Cells[lrenglon, lcolumna++].value = lrenglon  ; //Folio Cargo
                string sfecha = row["cfecha"].ToString().Trim();
                sfecha = sfecha.Substring(6, 2).PadLeft(2, '0') + "/" + sfecha.Substring(4, 2).PadLeft(2, '0') + "/" + sfecha.Substring(0, 4);
                sheet.Cells[lrenglon, lcolumna++].value = sfecha; //Folio Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["cfolio"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["crazonso01"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["ccodigop01"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["ccodaltern"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["cnombrep01"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["cnomaltern"].ToString().Trim(); //C
                sheet.Cells[lrenglon, lcolumna++].value = row["cunidades"].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row["cunidade03"].ToString().Trim(); //pendiente de facturar
                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;
        }
        public void mReporteFacturaAbono(string mEmpresa)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoFacturaAbono(sheet,mEmpresa, "", lrenglon);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosFacturaAbono.Rows)
            {
                lcolumna = 1;
                sheet.Cells[lrenglon, lcolumna++].value = row[0].ToString().Trim(); //concepto factura
                sheet.Cells[lrenglon, lcolumna++].value = row[1].ToString().Trim(); //folio factura
                sheet.Cells[lrenglon, lcolumna++].value = row[2].ToString().Trim(); //serie factura
                sheet.Cells[lrenglon, lcolumna++].value = row[3].ToString().Trim(); //total factura
                sheet.Cells[lrenglon, lcolumna++].value = row[4].ToString().Trim(); //referencia factura
                sheet.Cells[lrenglon, lcolumna++].value = row[5].ToString().Trim(); //pendiente factura
                sheet.Cells[lrenglon, lcolumna++].value = row[6].ToString().Trim(); //serie abono
                sheet.Cells[lrenglon, lcolumna++].value = row[7].ToString().Trim(); //folio abono

                string lconceptopago = row[9].ToString().Trim().ToUpper();
                sheet.Cells[lrenglon, 9].value = "";
                sheet.Cells[lrenglon, 10].value = "";
                sheet.Cells[lrenglon, 11].value = "";
                sheet.Cells[lrenglon, 12].value = "";
                sheet.Cells[lrenglon, 13].value = "";
                sheet.Cells[lrenglon, 14].value = "";
                sheet.Cells[lrenglon, 15].value = "";
                
                /*
                Efectivo

                   BBV

IXE

Transferencia

Cheque

NC

Devolución
                 */
                if (lconceptopago.IndexOf("Efectivo")!=-1)
                    sheet.Cells[lrenglon, 9].value = row[8].ToString().Trim();
                else
                    if (lconceptopago.IndexOf("BBV") != -1)
                        sheet.Cells[lrenglon, 10].value = row[8].ToString().Trim();
                    else
                        if (lconceptopago.IndexOf("IXE") != -1)
                            sheet.Cells[lrenglon, 11].value = row[8].ToString().Trim();
                        else
                            if (lconceptopago.IndexOf("Transferencia") != -1)
                            sheet.Cells[lrenglon, 12].value = row[8].ToString().Trim();
                        else
                            if (lconceptopago.IndexOf("Cheque") != -1)
                                sheet.Cells[lrenglon, 13].value = row[8].ToString().Trim();
                            else
                                if (lconceptopago.IndexOf("NC") != -1)
                                    sheet.Cells[lrenglon, 14].value = row[8].ToString().Trim();
                if (lconceptopago.IndexOf("Devolución") != -1)
                                    sheet.Cells[lrenglon, 15].value = row[8].ToString().Trim();
                                else
                                    sheet.Cells[lrenglon, 16].value = row[8].ToString().Trim();
                sheet.get_Range("I" + lrenglon.ToString(), "O" + lrenglon.ToString()).Style = "Currency";
                lrenglon++;


            }
            TotalizarReporteFacturaAbono(sheet,lrenglon);
            sheet.Cells.EntireColumn.AutoFit();
            return;



        }
        public List<RegConcepto> mCargarConceptos(string mEmpresa)
        {
             List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            if (mEmpresa.IndexOf("\\") != -1)
            {
                OleDbConnection lconexion = new OleDbConnection();

                lconexion = mAbrirConexionOrigen(mEmpresa);
                
                if (lconexion != null)
                {
                    OleDbCommand lsql = new OleDbCommand("select cidconce01,ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = 4", lconexion);
                    OleDbDataReader lreader;
                    lreader = lsql.ExecuteReader();
                    _RegFacturas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegConcepto lRegConcepto = new RegConcepto();
                            lRegConcepto.Codigo = lreader[1].ToString();
                            lRegConcepto.Nombre = lreader[2].ToString();
                            lRegConcepto.id = long.Parse(lreader[0].ToString());
                            _RegFacturas.Add(lRegConcepto);
                        }
                    }
                    lreader.Close();
                }
            }
            return _RegFacturas;
        }

        public void mCargarClasificaciones(string mEmpresa, int clasificacion)
        {

            int clasif = clasificacion + 24;
            //List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            _RegClasificaciones.Clear();
            if (mEmpresa.IndexOf("\\") != -1)
            {
                OleDbConnection lconexion = new OleDbConnection();

                lconexion = mAbrirConexionOrigen(mEmpresa);

                if (lconexion != null)
                {
                    OleDbCommand lsql = new OleDbCommand("select cidvalor01,ccodigov01,cvalorcl01 from mgw10020 where cidclasi01 = " + clasif, lconexion);
                    OleDbDataReader lreader = null ;
                    try
                    {
                        lreader = lsql.ExecuteReader();
                    }
                    catch (Exception eeee)
                    { 

                    }
                    _RegClasificaciones.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegConcepto lRegConcepto = new RegConcepto();
                            lRegConcepto.Codigo = lreader[1].ToString();
                            lRegConcepto.Nombre = lreader[2].ToString();
                            lRegConcepto.id = long.Parse(lreader[0].ToString());
                            _RegClasificaciones.Add(lRegConcepto);
                        }
                    }
                    lreader.Close();
                }
            }
            //return _RegClasificaciones;
        }

        public void mCargarClasificacionesComercial(string mEmpresa, int clasificacion)
        {


            int clasif = clasificacion + 24;
            //List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            _RegClasificaciones.Clear();
            if (mEmpresa.IndexOf("\\") != -1)
            {
                OleDbConnection lconexion = new OleDbConnection();

                lconexion = mAbrirConexionOrigen(mEmpresa);

                if (lconexion != null)
                {
                    OleDbCommand lsql = new OleDbCommand("select cidvalor01,ccodigov01,cvalorcl01 from mgw10020 where cidclasi01 = " + clasif, lconexion);
                    OleDbDataReader lreader = null;
                    try
                    {
                        lreader = lsql.ExecuteReader();
                    }
                    catch (Exception eeee)
                    {

                    }
                    _RegClasificaciones.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegConcepto lRegConcepto = new RegConcepto();
                            lRegConcepto.Codigo = lreader[1].ToString();
                            lRegConcepto.Nombre = lreader[2].ToString();
                            lRegConcepto.id = long.Parse(lreader[0].ToString());
                            _RegClasificaciones.Add(lRegConcepto);
                        }
                    }
                    lreader.Close();
                }
            }
        }

        public void mBorraElememento(RegConcepto clasif)
        {
            _RegClasificaciones.Remove(clasif);
        }

        public void mReportePedidoFactura(string mEmpresa, string lfechai, string lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoPedidoFactura(sheet, mEmpresa, "", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            foreach (DataRow row in DatosFacturaAbono.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total

                lcolumna = 1;
                sheet.Cells[lrenglon, lcolumna++].value = row[0].ToString().Trim(); //Fecha
                sheet.Cells[lrenglon, lcolumna++].value = row[1].ToString().Trim(); //serie pedidos
                sheet.Cells[lrenglon, lcolumna++].value = row[2].ToString().Trim(); //#pedidos
                sheet.Cells[lrenglon, lcolumna++].value = row[3].ToString().Trim(); //cliente
                sheet.Cells[lrenglon, lcolumna++].value = row[4].ToString().Trim(); //importe
                sheet.Cells[lrenglon, lcolumna++].value = row[5].ToString().Trim(); //pendiente de facturar

                sheet.Cells[lrenglon, lcolumna++].value = "'" + row[13].ToString().Trim(); //fecha de facturacion

                sheet.Cells[lrenglon, lcolumna++].value = row[6].ToString().Trim(); //serie factura
                sheet.Cells[lrenglon, lcolumna++].value = row[7].ToString().Trim(); //# de factura
                if (row[7].ToString().Trim() == "")
                    sheet.Cells[lrenglon, lcolumna++].value = ""; // cliente    
                else
                    sheet.Cells[lrenglon, lcolumna++].value = row[8].ToString().Trim(); // cliente
                sheet.Cells[lrenglon, lcolumna++].value = row[9].ToString().Trim(); // importe
                sheet.Cells[lrenglon, lcolumna++].value = row[10].ToString().Trim(); // impuesto
                sheet.Cells[lrenglon, lcolumna++].value = row[11].ToString().Trim(); // retencion
                sheet.Cells[lrenglon, lcolumna++].value = row[12].ToString().Trim(); // total


                
                sheet.get_Range("E" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";
                sheet.get_Range("K" + lrenglon.ToString(), "N" + lrenglon.ToString()).Style = "Currency";
                lrenglon++;


            }
            sheet.Cells.EntireColumn.AutoFit();
            return;



        }

        public string mRegresarCatalogoValido(int tipo,  string codigo, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string regresa = "";
            lconexion = mAbrirConexionOrigen(mEmpresa);
            OleDbCommand lsql = new OleDbCommand();
            if (lconexion != null)
            {
                if (tipo == 2) // proveedores
                    lsql.CommandText = "select cidclien01,crazonso01 from mgw10002 where ccodigoc01 = '" + codigo + "' and ctipocli01 >= 1";
                lsql.Connection = lconexion;
                OleDbDataReader lreader;
                lreader = lsql.ExecuteReader();
                
                if (lreader.HasRows)
                {
                    lreader.Read();
                    {
                        if (lreader[0].ToString() != "")
                        {
                            regresa = lreader[1].ToString();
                        }
                    }
                }
                lreader.Close();
                
            }
            return regresa;

        }
        public void mReporteForrajeraComercial(string mEmpresa, DateTime lfechai, DateTime lfechaf)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            configuracionencabezadoForrajeraComercial(sheet, mEmpresa, "REPORTE DE MOVIMIENTOS POR CONCEPTO POR PRODUCTO", lrenglon, lfechai, lfechaf);

            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon = 6;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            string conceptoprevio = "";
            decimal lcantidad = 0;
            decimal lcosto = 0;
            decimal ltotal = 0;
            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                string concepto = row["CNOMBRECONCEPTO"].ToString().Trim();
                if (concepto != conceptoprevio)
                {
                    if (conceptoprevio != "")
                    {
                        // totales

                        sheet.Cells[lrenglon, 1].value = "Total del Concepto";
                        sheet.Cells[lrenglon, 3].value = lcantidad;
                        //sheet.Cells[lrenglon, lcolumna].numberformat = "0.00";
                        sheet.Cells[lrenglon, 4].value = lcosto; //Serie Cargo
                        sheet.Cells[lrenglon, 5].value = ltotal; //Fecha Cargo
                        sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).Style = "Comma";
                        sheet.get_Range("D" + lrenglon.ToString(), "E" + lrenglon.ToString()).Style = "Currency";
                        sheet.get_Range("A" + lrenglon.ToString(), "E" + lrenglon.ToString()).Font.Bold = true;
                        lrenglon += 2;

                    }
                    // imprimir titulo de concepto
                    sheet.Cells[++lrenglon, 1].value = "Concepto: " + concepto;

                    sheet.get_Range("A" + lrenglon.ToString(), "A" + lrenglon.ToString()).Font.Bold = true;



                    lrenglon += 2;
                    conceptoprevio = concepto;
                    lcantidad = 0;
                    lcosto = 0;
                    ltotal = 0;
                }
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                //DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                //DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                //string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                //string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');


                //sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["CCODIGOPRODUCTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREPRODUCTO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["cantidad"].ToString().Trim();
                lcantidad += decimal.Parse(row["cantidad"].ToString().Trim());
                //sheet.Cells[lrenglon, lcolumna].numberformat = "0.00";
                sheet.Cells[lrenglon, lcolumna++].value = row["costo"].ToString().Trim(); //Serie Cargo
                lcosto += decimal.Parse(row["costo"].ToString().Trim());

                sheet.Cells[lrenglon, lcolumna++].value = row["TOTAL"].ToString().Trim(); //Fecha Cargo
                ltotal += decimal.Parse(row["TOTAL"].ToString().Trim());
                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C

                sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).Style = "Comma";
                sheet.get_Range("D" + lrenglon.ToString(), "E" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }
            //sheet.Cells.EntireColumn.AutoFit();
            sheet.get_Range("A" + lrenglon.ToString(), "A" + lrenglon.ToString()).EntireColumn.ColumnWidth = 25;
            sheet.get_Range("B" + lrenglon.ToString(), "B" + lrenglon.ToString()).EntireColumn.ColumnWidth = 55;
            sheet.get_Range("C" + lrenglon.ToString(), "C" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;
            sheet.get_Range("D" + lrenglon.ToString(), "D" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;
            sheet.get_Range("E" + lrenglon.ToString(), "E" + lrenglon.ToString()).EntireColumn.ColumnWidth = 20;

            return;
        }
        private void configuracionencabezadoForrajeraComercial(MyExcel.Worksheet sheet, string Empresa, string texto, int lrenglon, DateTime lfecha1, DateTime lfecha2)
        {
            int lcolumna = 1;
            //EncabezadoEmpresa(sheet, Empresa, texto);

            sheet.Cells[1, 5].value = texto;
            string fecha2 = lfecha1.Day.ToString().PadLeft(2, '0') + "/" + lfecha1.Month.ToString().PadLeft(2, '0') + "/" + lfecha1.Year.ToString().PadLeft(4, '0');
            string fecha3 = lfecha2.Day.ToString().PadLeft(2, '0') + "/" + lfecha2.Month.ToString().PadLeft(2, '0') + "/" + lfecha2.Year.ToString().PadLeft(4, '0');



            sheet.Cells[2, 5].value = "Fecha del " + fecha2 + " al " + fecha3;


            lrenglon += 4;
            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Interior.Color = Color.Blue;

            sheet.get_Range("A" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Font.Color = Color.White;

            sheet.Cells[lrenglon, lcolumna++].value = "Codigo Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Nombre Producto";
            sheet.Cells[lrenglon, lcolumna++].value = "Cantidad";
            sheet.Cells[lrenglon, lcolumna++].value = "Costo Unitario";
            sheet.Cells[lrenglon, lcolumna++].value = "Total";

            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "E" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
        }
    }

}
