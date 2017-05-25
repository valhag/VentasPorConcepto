using System;
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

    
    class Class1
    {
        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public OleDbConnection _conexion;
        private DataTable DatosFacturaAbono = null;

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

        public void mReporteInventarioCapas(string mEmpresa, string lfechai, string lfechaf)
        {

            mConfigurarObjetosImpresion();


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
              almacen = b.almacen
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

    }

}
