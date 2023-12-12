using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using LibreriaDoctos;
using MyExcel = Microsoft.Office.Interop.Excel;


namespace VentasPorConcepto
{
    public partial class FacturaPedido : ReporteBase
    {

        DataTable DatosReporte = null;
        DataTable DatosResumen = null;

        public FacturaPedido()
        {
            InitializeComponent();
        }

        private void FacturaPedido_Load(object sender, EventArgs e)
        {
            this.Text = " Reporte Facturas y Pedidos " + this.ProductVersion;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            mTraerInformacionComercial(empresasComercial1.aliasbdd);

        }

        public void mTraerInformacionComercial(string mEmpresa)

        {

            //MessageBox.Show(empresasComercial1.micombo.Text.ToString());
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


            StringBuilder lquery = new StringBuilder();

            DateTime lfecha1 = dateTimePicker1.Value;
            string sfecha1 = lfecha1.Year.ToString() + lfecha1.Month.ToString().PadLeft(2, '0') + lfecha1.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');



            lquery.Append(" select d.cfecha, a.CNOMBREAGENTE, d.CFOLIO as factura");
            lquery.Append(" , isnull(ped.cfolio, 0) as pedido");
            lquery.Append(" ,c.CRAZONSOCIAL, dom.CESTADO, p.CNOMBREPRODUCTO,");
            lquery.Append(" m.CUNIDADESCAPTURADAS, m.COBSERVAMOV, m.CPRECIOCAPTURADO, m.CNETO, m.CIMPUESTO1, m.ctotal, d.CPENDIENTE");
            lquery.Append(" , m.CUNIDADESORIGEN, u.CNOMBREUNIDAD, c.ccodigocliente, d.ctotal ctotaldocumento");
            lquery.Append(" from admDocumentos d");
            lquery.Append(" join admAgentes a on d.CIDAGENTE = a.CIDAGENTE");
            lquery.Append(" join admClientes c on c.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR");
            lquery.Append(" join admDomicilios dom on dom.CIDCATALOGO = d.CIDDOCUMENTO and dom.CTIPOCATALOGO = 3 and dom.CTIPODIRECCION = 1");
            lquery.Append(" join admMovimientos m on m.CIDDOCUMENTO = d.CIDDOCUMENTO");
            lquery.Append(" join admUnidadesMedidaPeso u on u.CIDUNIDAD = m.CIDUNIDAD");
            lquery.Append(" join admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO");
            lquery.Append(" left join admDocumentos ped on ped.CIDDOCUMENTOORIGEN = d.CIDDOCUMENTO");
            lquery.Append(" where m.CUNIDADESPENDIENTES <> 0");
            lquery.Append(" and d.CIDDOCUMENTODE = 4");
            lquery.Append(" and d.ccancelado = 0");

            lquery.Append(" AND D.Cfecha >= @fecha1");
            lquery.Append(" AND D.Cfecha <= @fecha2");
            lquery.Append(" ORDER BY d.cfolio, d.cfecha");
            DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, _conexion1);
            mySqlDataAdapter.SelectCommand.Parameters.AddWithValue("@fecha1", sfecha1);
            mySqlDataAdapter.SelectCommand.Parameters.AddWithValue("@fecha2", sfecha2);
            mySqlDataAdapter.Fill(ds);


            DatosReporte = ds.Tables[0];


            _conexion1.Close();

            mReporteComercial(empresasComercial1.aliasbdd);


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

        public void mReporteComercial(string mEmpresa)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];




            sheet.Cells[1, 1].value = "Comercial i";
            sheet.Cells[1, 5].value = empresasComercial1.micombo.Text.ToString();




            //if (radioButton1.Checked == true)
            sheet.Cells[2, 5].value = "REPORTE DE VENTAS";
            //else
            //  sheet.Cells[2, 5].value = "A / R Aging Summary";



            sheet.Cells[2, 11].value = System.DateTime.Today;



            sheet.Cells[3, 5].value = "del dia :" + dateTimePicker1.Value.Day.ToString().PadLeft(2, '0') + "/" + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + "/" + dateTimePicker1.Value.Year.ToString();
            sheet.Cells[4, 5].value = "al dia :" + dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + "/" + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + "/" + dateTimePicker2.Value.Year.ToString();




            mEncabezadoCelda(sheet, "A", "A", 5, 0, 10, "Fecha", false);
            mEncabezadoCelda(sheet, "B", "B", 5, 0, 10, "Ejecutivo", false);
            mEncabezadoCelda(sheet, "C", "C", 5, 0, 10, "No. Pedido", false);
            mEncabezadoCelda(sheet, "D", "D", 5, 0, 10, "No. Factura", false);

            mEncabezadoCelda(sheet, "E", "E", 5, 0, 10, "Codigo Cliente", false);

            mEncabezadoCelda(sheet, "F", "F", 5, 0, 10, "Cliente", false);
            mEncabezadoCelda(sheet, "G", "G", 5, 0, 10, "Estado", false);
            mEncabezadoCelda(sheet, "H", "H", 5, 0, 10, "Producto", false);
            mEncabezadoCelda(sheet, "I", "I", 5, 0, 10, "Cantidad", false);

            mEncabezadoCelda(sheet, "J", "J", 5, 0, 10, "Unidad de Medida", false);

            mEncabezadoCelda(sheet, "K", "K", 5, 0, 10, "Descripcion", false);
            mEncabezadoCelda(sheet, "L", "L", 5, 0, 10, "Precio Unit", false);
            mEncabezadoCelda(sheet, "M", "M", 5, 0, 10, "Monto", false);
            mEncabezadoCelda(sheet, "N", "N", 5, 0, 10, "IVA", false);
            mEncabezadoCelda(sheet, "O", "O", 5, 0, 10, "Total", false);

            mEncabezadoCelda(sheet, "P", "P", 5, 0, 10, "Total Documento", false);

            mEncabezadoCelda(sheet, "Q", "Q", 5, 0, 10, "Saldo Vencido", false);


            /*
            sheet.get_Range("H" + 5.ToString(), "I" + 5.ToString()).Merge();

            sheet.get_Range("A" + 5.ToString(), "I" +
            5.ToString()).Interior.Color = Color.LightBlue;
            */

            lrenglon = 6;
            //mResetearrTotales();


            


            foreach (DataRow row in DatosReporte.Rows)
            {

                //cfecha CNOMBREAGENTE   factura pedido  CRAZONSOCIAL CESTADO CNOMBREPRODUCTO CUNIDADESCAPTURADAS COBSERVAMOV CPRECIOCAPTURADO    
                //CNETO CIMPUESTO1  ctotal CPENDIENTE  CUNIDADESORIGEN

                //2017 - 12 - 06 00:00:00.000(Ninguno)                                                       1   0   NORA DE LA CRUZ MENDOZA Jalisco PRUEBA DE PRODUCTO  1   NULL    8000    8000    1280    9280    0   0

                int lcolumna=1;
                DateTime dfecha = DateTime.Parse(row["CFECHA"].ToString().Trim());
                string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2;

                sheet.Cells[lrenglon, lcolumna++].value = row["Cnombreagente"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["pedido"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["factura"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = "'" + row["ccodigocliente"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["crazonsocial"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cestado"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cnombreproducto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cunidadescapturadas"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["cnombreunidad"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["cobservamov"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cpreciocapturado"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cneto"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cimpuesto1"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["ctotal"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["ctotaldocumento"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["cpendiente"].ToString().Trim();
                



                //sheet.get_Range("D" + "6".ToString(), "I" + lrenglon.ToString()).Style = "CURRENCY";
                //sheet.get_Range("G" + lrenglon.ToString(), "G" + lrenglon.ToString()).NumberFormat = "0.00%";


                /*
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
                */



                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                //sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }



            //            sheet.Cells.EntireColumn.AutoFit();

            sheet.get_Range("L" + "6".ToString(), "Q" + lrenglon.ToString()).Style = "Currency";


            return;
        }

    }


}
