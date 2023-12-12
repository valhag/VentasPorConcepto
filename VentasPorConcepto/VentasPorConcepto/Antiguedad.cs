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
    public partial class Antiguedad : ReporteBase
    {

        DataTable DatosReporte = null;

        public Antiguedad()
        {
            InitializeComponent();
        }

        private void OnComboChange(object sender, EventArgs e)
        {
            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();
           // codigocatalogocomercial1.lrn.lbd.cadenaconexion = Cadenaconexion;
           // codigocatalogocomercial1.lrn = lrn;

        }

        private void Antiguedad_Load(object sender, EventArgs e)
        {

            codigocatalogocomercial3.Visible = false;
            codigocatalogocomercial4.Visible = false;

            codigocatalogocomercial1.mSetLabelText("Cliente Inicial");
            codigocatalogocomercial2.mSetLabelText("Cliente Final");

            codigocatalogocomercial3.mSetLabelText("Prov Inicial");
            codigocatalogocomercial4.mSetLabelText("Prov Final");

            codigocatalogocomercial1.TextLeave += new EventHandler(OnTextLeave1);
            codigocatalogocomercial2.TextLeave += new EventHandler(OnTextLeave2);


            codigocatalogocomercial3.TextLeave += new EventHandler(OnTextLeave3);
            codigocatalogocomercial4.TextLeave += new EventHandler(OnTextLeave4);

            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
        }

        public void OnTextLeave1(object sender, EventArgs e)
        {
            if (codigocatalogocomercial1.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(1, codigocatalogocomercial1.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial1.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Cliente no Existe");
                    codigocatalogocomercial1.mSetFocus();
                }
            }
        }

        public void OnTextLeave2(object sender, EventArgs e)
        {
            if (codigocatalogocomercial2.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(1, codigocatalogocomercial2.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial2.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Cliente no Existe");
                    codigocatalogocomercial2.mSetFocus();
                }
            }
        }

        public void OnTextLeave3(object sender, EventArgs e)
        {
            if (codigocatalogocomercial3.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(3, codigocatalogocomercial3.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial1.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Proveedor no Existe");
                    codigocatalogocomercial3.mSetFocus();
                }
            }
        }

        public void OnTextLeave4(object sender, EventArgs e)
        {
            if (codigocatalogocomercial4.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(3, codigocatalogocomercial4.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial4.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Proveedor no Existe");
                    codigocatalogocomercial4.mSetFocus();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DateTime.Today > DateTime.Parse("2020/10/01"))
            {
                //MessageBox.Show ("La configuracion de adminpaq no es correcta");
              //  return ;
            }

            mTraerInformacionComercial(empresasComercial1.aliasbdd);
        }

        

//        public void mTraerInformacionComercial(StringBuilder lquery, string mEmpresa)
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

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');


            lquery.Append("with cte");
            lquery.Append(" as");
            lquery.Append(" (");
            lquery.Append(" select c.CDIASREVISION,c.CDIAPAGO, d.CFECHAVENCIMIENTO, d.cfecha,d.CSERIEDOCUMENTO,c.CCODIGOCLIENTE,d.CRAZONSOCIAL,d.cpendiente, d.cfolio, datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') dias");
            lquery.Append(" ,case when datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') < 0 then cpendiente else 0 end actual");
            lquery.Append("   ,case when datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') between 1 and 30 then cpendiente else 0 end periodo130");
            lquery.Append("     ,case when datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') between 31 and 60 then cpendiente else 0 end periodo3160");
            lquery.Append("       ,case when datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') between 61 and 90 then cpendiente else 0 end periodo6190");
            lquery.Append("         ,case when datediff(day, d.CFECHAVENCIMIENTO, '" + sfecha1 + "') > 90 then cpendiente else 0 end periodo91");
            lquery.Append("           , SUM(D.CPENDIENTE) OVER");
            lquery.Append("                    (PARTITION BY c.CIDCLIENTEPROVEEDOR) AS total");
            lquery.Append(", co.CNOMBRECONCEPTO ");
            lquery.Append(", case when month(d.cfechavencimiento) = 1 then 'JAN'");
            lquery.Append(" when month(d.cfechavencimiento) = 2 then 'FEB'");
            lquery.Append(" when month(d.cfechavencimiento) = 3 then 'MAR'");
            lquery.Append(" when month(d.cfechavencimiento) = 4 then 'APR'");
            lquery.Append(" when month(d.cfechavencimiento) = 5 then 'MAY'");
            lquery.Append(" when month(d.cfechavencimiento) = 6 then 'JUN'");
            lquery.Append(" when month(d.cfechavencimiento) = 7 then 'JUL'");
            lquery.Append(" when month(d.cfechavencimiento) = 8 then 'AUG'");
            lquery.Append(" when month(d.cfechavencimiento) = 9 then 'SEP'");
            lquery.Append(" when month(d.cfechavencimiento) = 10 then 'OCT'");
            lquery.Append(" when month(d.cfechavencimiento) = 11 then 'NOV'");
            lquery.Append(" when month(d.cfechavencimiento) = 12 then 'DEC'");
            lquery.Append(" END mesvenc");
            lquery.Append(",case when month(d.cfecha) = 1 then 'JAN'");
            lquery.Append(" when month(d.cfecha) = 2 then 'FEB'");
            lquery.Append(" when month(d.cfecha) = 3 then 'MAR'");
            lquery.Append(" when month(d.cfecha) = 4 then 'APR'");
            lquery.Append(" when month(d.cfecha) = 5 then 'MAY'");
            lquery.Append(" when month(d.cfecha) = 6 then 'JUN'");
            lquery.Append(" when month(d.cfecha) = 7 then 'JUL'");
            lquery.Append(" when month(d.cfecha) = 8 then 'AUG'");
            lquery.Append(" when month(d.cfecha) = 9 then 'SEP'");
            lquery.Append(" when month(d.cfecha) = 10 then 'OCT'");
            lquery.Append(" when month(d.cfecha) = 11 then 'NOV'");
            lquery.Append(" when month(d.cfecha) = 12 then 'DEC'");
            lquery.Append("END MES");
            lquery.Append(" from admDocumentos d");
            lquery.Append(" join admClientes c on c.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR");
            lquery.Append(" join admConceptos co on co.CIDCONCEPTODOCUMENTO = d.CIDCONCEPTODOCUMENTO");
            
            if (radioButton1.Checked == true)
            lquery.Append(" where d.CIDDOCUMENTODE = 4");
            else
                lquery.Append(" where d.CIDDOCUMENTODE = 19");

            lquery.Append(" and d.cpendiente > 0");

            if (codigocatalogocomercial1.mGetCodigo() != "" && codigocatalogocomercial2.mGetCodigo() != "")
                lquery.Append("and c.CCODIGOCLIENTE between '" + codigocatalogocomercial1.mGetCodigo() + "' and '" + codigocatalogocomercial2.mGetCodigo() + "'");
            lquery.Append(" )");
            lquery.Append(" ");
            lquery.Append(" select *");
            lquery.Append(" ,sum(actual) over(partition by crazonsocial) sumactual");
            lquery.Append(" ,sum(periodo130) over(partition by crazonsocial) sumperiodo130");
            lquery.Append(" ,sum(periodo3160) over(partition by crazonsocial) sumperiodo3160");
            lquery.Append(" ,sum(periodo6190) over(partition by crazonsocial) sumperiodo6190");
            lquery.Append(" ,sum(periodo91) over(partition by crazonsocial) sumperiodo91");
            lquery.Append(" ,sum(actual) over() sumactualgtotal");
            lquery.Append(" ,sum(periodo130) over() sumperiodo130gtotal");
            lquery.Append(" ,sum(periodo3160) over() sumperiodo3160gtotal");
            lquery.Append(" ,sum(periodo6190) over() sumperiodo6190gtotal");
            lquery.Append(" ,sum(periodo91) over() sumperiodo91gtotal");
            lquery.Append(" ,sum(cpendiente) over() sumtotalgtotal");

            lquery.Append(" from cte");
            lquery.Append(" ORDER BY ccodigocliente, CFECHAVENCIMIENTO");



            DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, _conexion1);


            mySqlDataAdapter.Fill(ds);

             
            DatosReporte = ds.Tables[0];
            //if (ds.Tables.Count > 1)
            //    DatosDetalle = ds.Tables[1];

            mReporteComercial( empresasComercial1.aliasbdd);


            _conexion1.Close();

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

        private void mEncabezadoCelda(MyExcel.Worksheet sheet, string inicio, string fin, int lrenglon, int lagregarenglon, int tamano, string texto, Boolean Bold = true)
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
        public void mReporteComercial(string mEmpresa)
        {
            MyExcel.Workbook newWorkbook = mIniciarExcel();
            int lrenglon = 1;
            int lrengloninicial = 1;
            int lrengloniniciaconcepto = 1;
            int lrenglontempo = 1;
            MyExcel.Worksheet sheet = newWorkbook.Sheets[1];

            //configuracionencabezadoRemisionesComercial(sheet, mEmpresa, "Facturas y Pedidos", lrenglon, lfechai, lfechaf);



            sheet.Cells[1, 1].value = "CONTPAQ i";
            sheet.Cells[1, 5].value = empresasComercial1.micombo.Text.ToString();




            if (radioButton1.Checked == true)
            sheet.Cells[2, 5].value = "A / R Aging Summary";
            else
                sheet.Cells[2, 5].value = "A / R Aging Summary";

            string mes3 = "";
            switch (dateTimePicker1.Value.Month)
            {
                case 1: mes3 = "JANUARY"; break;
                case 2: mes3 = "FEBRUARY"; break;
                case 3: mes3 = "MARCH"; break;
                case 4: mes3 = "APRIL"; break;
                case 5: mes3 = "MAY"; break;
                case 6: mes3 = "JUNE"; break;
                case 7: mes3 = "JULY"; break;
                case 8: mes3 = "AUGUST"; break;
                case 9: mes3 = "SEPTEMBER"; break;
                case 10: mes3 = "OCTOBER"; break;
                case 11: mes3 = "NOVEMBER"; break;
                case 12: mes3 = "DECEMBER"; break;

            }


            sheet.Cells[2, 11].value =   System.DateTime.Today;

            string mes1 = "";
            switch (dateTimePicker1.Value.Month) {
                case 1:mes1 = "JANUARY"; break;
                case 2: mes1 = "FEBRUARY"; break;
                case 3: mes1 = "MARCH"; break;
                case 4: mes1 = "APRIL"; break;
                case 5: mes1 = "MAY"; break;
                case 6: mes1 = "JUNE"; break;
                case 7: mes1 = "JULY"; break;
                case 8: mes1 = "AUGUST"; break;
                case 9: mes1 = "SEPTEMBER"; break;
                case 10: mes1 = "OCTOBER"; break;
                case 11: mes1 = "NOVEMBER"; break;
                case 12: mes1 = "DECEMBER"; break;
               
            }

            sheet.Cells[3, 5].value = "As of:" + mes1 + " " + dateTimePicker1.Value.Day + "," + dateTimePicker1.Value.Year.ToString();



            //Comercial Royal SYL Ventures México, S de R.L.de C.V.Hoja:      1
            //Moneda: Dólar Americano                                                                                       17 / AGO / 2020

            //                   Fecha de Corte: 17 / AGO / 2020


            //        Saldo Numero  Días Actual  1 - 30    31 - 60   61 - 90   91.00
            //Vencimiento Fecha   Serie Total   Factura Días    Días Días    Días o más Concepto

            //mEncabezadoCelda(MyExcel.Worksheet sheet, string inicio, string fin, int lrenglon, int lagregarenglon, int tamano, string texto, Boolean Bold = true)


            mEncabezadoCelda(sheet, "A", "A", 5, 0, 8, "Due Date", false);
            mEncabezadoCelda(sheet, "B", "B", 5, 0, 8, "Date", false);
            mEncabezadoCelda(sheet, "C", "C", 5, 0, 8, "Serie", false);

            mEncabezadoCelda(sheet, "D", "D", 5, 0, 8, "Invoice", false);
            mEncabezadoCelda(sheet, "E", "E", 5, 0, 8, "Due Days", false);
            mEncabezadoCelda(sheet, "F", "F", 5, 0, 8, "TOTAL", false);
            mEncabezadoCelda(sheet, "G", "G", 5, 0, 8, "Current", false);
            mEncabezadoCelda(sheet, "H", "H", 5, 0, 8, "01-30", false);
            mEncabezadoCelda(sheet, "I", "I", 5, 0, 8, "31-60", false);
            mEncabezadoCelda(sheet, "J", "J", 5, 0, 8, "61-90", false);
            mEncabezadoCelda(sheet, "K", "K", 5, 0, 8, "91", false);
            mEncabezadoCelda(sheet, "L", "L", 5, 0, 8, "Concept", false);

            sheet.get_Range("A" + 5.ToString(), "L" +
            5.ToString()).Interior.Color = Color.LightGray;



            //    Date Serie   Invoice Due Days TOTAL   Current 01 - 30 31 - 60   61 - 90   91  Concept


            /*mEncabezadoCelda(sheet, "D", "D", 5, 0, 8, "Balance", false);
            mEncabezadoCelda(sheet, "E", "E", 5, 0, 8, "Number", false);
            mEncabezadoCelda(sheet, "F", "F", 5, 0, 8, "Days", false);
            mEncabezadoCelda(sheet, "G", "G", 5, 0, 8, "Current", false);
            mEncabezadoCelda(sheet, "H", "H", 5, 0, 8, "1-30", false);
            mEncabezadoCelda(sheet, "I", "I", 5, 0, 8, "31-60", false);
            mEncabezadoCelda(sheet, "J", "J", 5, 0, 8, "61-90", false);
            mEncabezadoCelda(sheet, "K", "K", 5, 0, 8, "91", false);


            mEncabezadoCelda(sheet, "A", "A", 6, 0, 8, "Due Date", false);
            mEncabezadoCelda(sheet, "B", "B", 6, 0, 8, "Date", false);
            mEncabezadoCelda(sheet, "C", "C", 6, 0, 8, "Serie", false);

            mEncabezadoCelda(sheet, "D", "D", 6, 0, 8, "Total", false);
            if (radioButton1.Checked == true)
                mEncabezadoCelda(sheet, "E", "E", 6, 0, 8, "Invoice", false);
            else
                mEncabezadoCelda(sheet, "E", "E", 6, 0, 8, "Purchase", false);

            mEncabezadoCelda(sheet, "F", "F", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "G", "G", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "H", "H", 6, 0, 8, "Days", false);
            mEncabezadoCelda(sheet, "I", "I", 6, 0, 8, "Days", false);
            mEncabezadoCelda(sheet, "J", "J", 6, 0, 8, "Days", false);
            mEncabezadoCelda(sheet, "K", "K", 6, 0, 8, "Days", false);
            mEncabezadoCelda(sheet, "L", "L", 6, 0, 8, "Concept", false);

    **/
            //mEncabezadoCelda(sheet, "A", "A", 6, 0, 8, "Due", false);
            //mEncabezadoCelda(sheet, "B", "B", 6, 0, 8, "Date", false);
            //mEncabezadoCelda(sheet, "C", "C", 6, 0, 8, "Serie", false);

            /*mEncabezadoCelda(sheet, "A", "A", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "B", "B", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "C", "C", 6, 0, 8, "", false);
                                              
            mEncabezadoCelda(sheet, "D", "D", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "E", "E", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "F", "F", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "G", "G", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "H", "H", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "I", "I", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "J", "J", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "K", "K", 6, 0, 8, "", false);
            mEncabezadoCelda(sheet, "L", "L", 6, 0, 8, "", false);
            */

            lrenglon = 6;
            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;

            string sumtotal ="";

            string sumactual ="";
            string sumperiodo130 = "";
            string sumperiodo3160 = "";
            string sumperiodo6190 = "";
            string sumperiodo91 = "";

            string sumgactual = "";
            string sumgperiodo130 = "";
            string sumgperiodo3160 = "";
            string sumgperiodo6190 = "";
            string sumgperiodo91 = "";
            string sumgtotal = "";



            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                if (lcliente != row["CCODIGOCLIENTE"].ToString().Trim())
                {
                    if (lrenglon != 7)
                    {

                        sheet.get_Range( "A"+ lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;

                        sheet.Cells[++lrenglon, 6].value = sumtotal;

                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


                        sheet.Cells[lrenglon, 7].value = sumactual;

                        sheet.Cells[lrenglon, 8].value = sumperiodo130;

                        sheet.Cells[lrenglon, 9].value = sumperiodo3160;
                        sheet.Cells[lrenglon, 10].value = sumperiodo6190;
                        sheet.get_Range("G" + lrenglon.ToString(), "K" + lrenglon.ToString()).Style = "Comma";

                        sheet.Cells[lrenglon++, 11].value = sumperiodo91;
                        sheet.get_Range("A" + (lrenglon-1).ToString(), "K" + (lrenglon-1).ToString()).Font.Size = 8;
                        

                        //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                        //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                        //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).HorizontalAlignment = MyExcel.XlHAlign.xlHAlignCenter;
                        //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).VerticalAlignment = MyExcel.XlHAlign.xlHAlignCenter;

                    }
                    lcliente = row["CCODIGOCLIENTE"].ToString().Trim();

                    //sheet.Cells[++lrenglon, 4].value = row["total"].ToString().Trim();
                    sumactual = row["sumactual"].ToString().Trim();

                    sumtotal = row["total"].ToString().Trim();

                    sumperiodo130 = row["sumperiodo130"].ToString().Trim();

                    sumperiodo3160= row["sumperiodo3160"].ToString().Trim();
                    sumperiodo6190 = row["sumperiodo6190"].ToString().Trim();

                    sumperiodo91 = row["sumperiodo91"].ToString().Trim();


                    sumgactual = row["sumactualgtotal"].ToString().Trim();

                    sumgtotal = row["sumtotalgtotal"].ToString().Trim();

                    sumgperiodo130 = row["sumperiodo130gtotal"].ToString().Trim();

                    sumgperiodo3160 = row["sumperiodo3160gtotal"].ToString().Trim();
                    sumgperiodo6190 = row["sumperiodo6190gtotal"].ToString().Trim();

                    sumgperiodo91 = row["sumperiodo91gtotal"].ToString().Trim();



                    sheet.get_Range("A" + (lrenglon).ToString(), "K" + (lrenglon).ToString()).Font.Size = 8;
                    sheet.get_Range("A" + (lrenglon).ToString(), "A" + lrenglon).Font.Bold = true;

                    if (radioButton1.Checked == true)
                        sheet.Cells[lrenglon++, lcolumna].value = "Customer:" + row["CCODIGOCLIENTE"].ToString().Trim();
                    else
                        sheet.Cells[lrenglon++, lcolumna].value = "Supplier:" + row["CCODIGOCLIENTE"].ToString().Trim();
                    

                    int diapago = int.Parse(row["CDIAPAGO"].ToString().Trim());
                    //diapago = 3;
                    string otro="";
                    var b = 1;
                    b = diapago & 1;
                    if (b != 0)
                        otro += "L,";
                    b = diapago & 2;
                    if (b != 0)
                        otro += "M,";
                    b = diapago & 4;
                    if (b != 0)
                        otro += "I,";
                    b = diapago & 8;

                    if (b != 0)
                        otro += "J,";
                    b = diapago & 16;

                    if (b != 0)
                        otro += "V,";
                    b = diapago & 32;

                    if (b != 0)
                        otro += "S,";
                    b = diapago & 64;

                    if (b != 0)
                        otro += "D,";

                    int diarevision = int.Parse(row["CDIASREVISION"].ToString().Trim());
                    //diapago = 3;
                    string otro1 = "";
                    b = diarevision & 1;
                    if (b != 0)
                        otro1 += "L,";
                    b = diarevision & 2;
                    if (b != 0)
                        otro1 += "M,";
                    b = diarevision & 4;
                    if (b != 0)
                        otro1 += "I,";
                    b = diarevision & 8;

                    if (b != 0)
                        otro1 += "J,";
                    b = diarevision & 16;

                    if (b != 0)
                        otro1 += "V,";
                    b = diarevision & 32;

                    if (b != 0)
                        otro1 += "S,";
                    b = diarevision & 64;

                    if (b != 0)
                        otro1 += "D,";






                    //if (diapago & 0)
                    otro = otro.Substring(0,otro.Length-1);

                    //sheet.Cells[lrenglon++, lcolumna].value = "Payment Days:" + otro;

                    //sheet.Cells[lrenglon++, lcolumna].value = "Revision Days:" + otro1;
                    sheet.Cells[lrenglon++, lcolumna].value = "Name:" + row["CRAZONSOCIAL"].ToString().Trim();

                    sheet.get_Range("A" + (lrenglon-2).ToString(), "A" + (lrenglon-1).ToString()).Font.Size = 8;
                    sheet.get_Range("A" + (lrenglon - 2).ToString(), "A" + (lrenglon - 1).ToString()).Font.Bold = true;


                    sheet.get_Range("A" + (lrenglon - 1).ToString(), "A" + (lrenglon - 2).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                    sheet.get_Range("A" + (lrenglon - 1).ToString(), "A" + (lrenglon-2).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                    sheet.get_Range("A" + (lrenglon - 1).ToString(), "A" + (lrenglon-2).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                    sheet.get_Range("A" + (lrenglon - 1).ToString(), "A" + (lrenglon - 2).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                    sheet.get_Range("A" + (lrenglon - 1).ToString(), "A" + (lrenglon - 2).ToString()).Borders[MyExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;


                    //sheet.get_Range(inicio + lrenglon, fin + (lrenglon + lagregarenglon)).Font.Bold = Bold;

                }

                lcolumna = 1;
                //sheet.Cells[lrenglon, lcolumna++].value = lrenglon; //Folio Cargo
                DateTime dfecha = DateTime.Parse(row["cfecha"].ToString().Trim());

                DateTime dfechav = DateTime.Parse(row["CFECHAVENCIMIENTO"].ToString().Trim());
                string fecha2 = dfecha.Day.ToString().PadLeft(2, '0') + "/" + row["MES"].ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                string fechav = dfechav.Day.ToString().PadLeft(2, '0') + "/" + row["MESvenc"].ToString().Trim() + "/" + dfechav.Year.ToString().PadLeft(4, '0');

                sheet.Cells[lrenglon, lcolumna++].value = "'" + fechav;
                sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2;

                sheet.get_Range("A" + (lrenglon).ToString(), "L" + lrenglon).Font.Size = 8;


                sheet.Cells[lrenglon, lcolumna++].value = row["CSERIEDOCUMENTO"].ToString().Trim();
                sheet.get_Range("G" + lrenglon.ToString(), "K" + lrenglon.ToString()).Style = "Comma";
                


                sheet.Cells[lrenglon, lcolumna++].value = row["CFOLIO"].ToString().Trim(); //Serie Cargo

                sheet.Cells[lrenglon, lcolumna++].value = row["DIAS"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["CPENDIENTE"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["ACTUAL"].ToString().Trim(); //Serie Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["PERIODO130"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["PERIODO3160"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["PERIODO6190"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["PERIODO91"].ToString().Trim(); //Fecha Cargo
                sheet.Cells[lrenglon, lcolumna++].value = row["cnombreconcepto"].ToString().Trim(); //Fecha Cargo

                sheet.get_Range("A" + (lrenglon).ToString(), "L" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "L" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "L" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "L" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "L" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                //sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }

            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;

            sheet.Cells[++lrenglon, 6].value = sumtotal;

            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "L" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


            sheet.Cells[lrenglon, 7].value = sumactual;

            sheet.Cells[lrenglon, 8].value = sumperiodo130;

            sheet.Cells[lrenglon, 9].value = sumperiodo3160;
            sheet.Cells[lrenglon, 10].value = sumperiodo6190;
            sheet.get_Range("G" + lrenglon.ToString(), "K" + lrenglon.ToString()).Style = "Comma";

            sheet.Cells[lrenglon++, 11].value = sumperiodo91;
            sheet.get_Range("A" + (lrenglon - 1).ToString(), "K" + (lrenglon - 1).ToString()).Font.Size = 8;


            lrenglon += 2;

            sheet.Cells[lrenglon, 5].value = "TOTAL";

            sheet.Cells[lrenglon, 6].value = sumgtotal;
            sheet.Cells[lrenglon, 7].value = sumgactual;
            sheet.Cells[lrenglon, 8].value = sumgperiodo130;
            sheet.Cells[lrenglon, 9].value = sumgperiodo3160;
            sheet.Cells[lrenglon, 10].value = sumgperiodo6190;
            sheet.Cells[lrenglon, 11].value = sumgperiodo91;
            //TOTAL Current 30 - ene  31 - 60   61 - 90   91
            //sheet.get_Range("E" + (lrenglon).ToString(), "K" + (lrenglon).ToString()).Font.Size = 8;
            sheet.get_Range("E" + (lrenglon).ToString(), "K" + (lrenglon).ToString()).Font.Bold = true;
            sheet.get_Range("E" + lrenglon.ToString(), "K" +
            lrenglon.ToString()).Interior.Color = Color.LightGray;
            sheet.get_Range("F" + lrenglon.ToString(), "K" + lrenglon.ToString()).Style = "Comma";






            //sheet.Cells.EntireColumn.AutoFit();
            return;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                codigocatalogocomercial3.Visible = true;
                codigocatalogocomercial4.Visible = true;
                codigocatalogocomercial1.Visible = false;
                codigocatalogocomercial2.Visible = false;

            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                codigocatalogocomercial3.Visible = false;
                codigocatalogocomercial4.Visible = false;
                codigocatalogocomercial1.Visible = true;
                codigocatalogocomercial2.Visible = true;

            }
        }
    }
}
