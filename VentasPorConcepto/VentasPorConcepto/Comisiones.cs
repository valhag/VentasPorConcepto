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
    public partial class Comisiones : ReporteBase
    {

        DataTable DatosReporte = null;
        DataTable DatosResumen = null;

        public Comisiones()
        {
            InitializeComponent();
        }

        private void OnComboChange(object sender, EventArgs e)
        {
            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();
         
        }

        private void Comisiones_Load(object sender, EventArgs e)
        {
            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
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

            DateTime lfecha1 = dateTimePicker1.Value;
            string sfecha1 = lfecha1.Year.ToString() + lfecha1.Month.ToString().PadLeft(2, '0') + lfecha1.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');



            lquery.Append(" SELECT *");
            lquery.Append(" , SUM(monto) over(partition by cnombreagente order by cnombreagente) summonto");
            lquery.Append(" , SUM(costo) over(partition by cnombreagente order by cnombreagente) sumcosto");
            lquery.Append(" , SUM(utilidad) over(partition by cnombreagente order by cnombreagente) sumutilidad");
            lquery.Append(" , SUM(comisionpropia) over(partition by cnombreagente order by cnombreagente) sumcomisionpropia");
            lquery.Append("  , SUM(comisioN2) over(partition by cnombreagente order by cnombreagente) sumcomision2");
            lquery.Append("  , SUM(monto) over() totalmonto");
            lquery.Append("  , SUM(costo) over() totalcosto");
            lquery.Append("  , SUM(utilidad) over() totalutilidad");
            lquery.Append("  , SUM(comisionpropia) over() totalcomisionpropia");
            lquery.Append("  , SUM(comision2) over() totalcomision2");
            lquery.Append("  , porcentaje1/100 porcentaje");
            lquery.Append(" into #temphva ");
            lquery.Append("  FROM");
            lquery.Append(" (");
            lquery.Append(" SELECT *, MONTO - COSTO UTILIDAD");
            lquery.Append(" , CASE WHEN ISNULL(COSTO, 0) = 0 or costo = 0 or monto = 0 THEN 100 else (monto - costo) * 100 / monto end porcentaje1");
            lquery.Append("          ,(CCOMISIONVENTAAGENTE / 100) *(monto - costo) as comisionpropia");
            //lquery.Append(" , case when CNOMBREAGENTE not like 'ENRIQUE%' THEN((monto - costo) * (0.3333)) ELSE 1 END comision2");
            lquery.Append("                 , case when x.cseriedocumento = 'O'  THEN((monto - costo) * (0.03)) ELSE case when x.cseriedocumento = 'E'  THEN((monto - costo) * (0.03))ELSE 0 END END comision2");
            //lquery.Append("                 , case when x.cseriedocumento <> 'E' and x.ccodigoagente = '2' THEN((monto - costo) * (0.3333)) ELSE 0 END comision3");
            lquery.Append(" FROM");
            lquery.Append(" (");
            lquery.Append(" select a.CNOMBREAGENTE, d.CSERIEDOCUMENTO + ltrim(str(d.CFOLIO)) FACTURA, c.CRAZONSOCIAL CLIENTE, d.cneto MONTO,");
            lquery.Append(" sum(m.CCOSTOESPECIFICO) COSTO, A.CCOMISIONVENTAAGENTE, d.cseriedocumento, a.ccodigoagente");
            lquery.Append(" from admDocumentos d");
            lquery.Append(" join admAgentes a");
            lquery.Append(" on d.CIDAGENTE = a.CIDAGENTE and d.ciddocumentode = 4 and d.ccancelado = 0 ");
            lquery.Append(" join admClientes c");
            lquery.Append(" on c.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR");
            lquery.Append(" join admMovimientos m");
            lquery.Append(" on m.CIDDOCUMENTO = d.CIDDOCUMENTO");
            lquery.Append(" where CNOMBREAGENTE <> '(Ninguno)'");
            lquery.Append(" AND D.Cfecha >= @fecha1");
            lquery.Append(" AND D.Cfecha <= @fecha2");

            //lquery.Append(" --and CNOMBREAGENTE like 'EXTERNO%'");
            lquery.Append(" group by a.CNOMBREAGENTE, d.CSERIEDOCUMENTO + ltrim(str(d.CFOLIO)), c.CRAZONSOCIAL, d.cneto, A.CCOMISIONCOBROAGENTE, A.CCOMISIONVENTAAGENTE,CSERIEDOCUMENTO, a.CCODIGOAGENTE");
            lquery.Append(" ) X");
            lquery.Append(" ) Y");
            lquery.Append(" ORDER BY CNOMBREAGENTE; select * from #temphva ORDER BY CNOMBREAGENTE");
            lquery.Append(" ; select cnombreagente, avg(sumcomisionpropia) as comision1, avg(sumcomision2) as comision2  from #temphva group by cnombreagente ORDER BY CNOMBREAGENTE");

            DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, _conexion1);
            mySqlDataAdapter.SelectCommand.Parameters.AddWithValue("@fecha1", sfecha1);
            mySqlDataAdapter.SelectCommand.Parameters.AddWithValue("@fecha2", sfecha2);

            //da = new SqlDataAdapter("SELECT * FROM annotations WHERE annotation LIKE @search",
            //                      _mssqlCon.connection);
            //        da.SelectCommand.Parameters.AddWithValue("@search", "%" + txtSearch.Text + "%");

            mySqlDataAdapter.Fill(ds);


            DatosReporte = ds.Tables[0];

            DatosResumen = ds.Tables[1];
            //if (ds.Tables.Count > 1)
            //    DatosDetalle = ds.Tables[1];

            mReporteComercial(empresasComercial1.aliasbdd);


            _conexion1.Close();

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
                sheet.Cells[2, 5].value = "Comisiones";
            //else
              //  sheet.Cells[2, 5].value = "A / R Aging Summary";

            

            sheet.Cells[2, 11].value = System.DateTime.Today;

            

            sheet.Cells[3, 5].value = "del dia :" + dateTimePicker1.Value.Day.ToString().PadLeft(2,'0') + "/" + dateTimePicker1.Value.Month.ToString().PadLeft(2, '0') + "/" + dateTimePicker1.Value.Year.ToString();
            sheet.Cells[4, 5].value = "al dia :" + dateTimePicker2.Value.Day.ToString().PadLeft(2, '0') + "/" + dateTimePicker2.Value.Month.ToString().PadLeft(2, '0') + "/" + dateTimePicker2.Value.Year.ToString();



            
            mEncabezadoCelda(sheet, "A", "A", 5, 0, 8, "NOMBRE DEL AGENTE ", false);
            mEncabezadoCelda(sheet, "B", "B", 5, 0, 8, "FACTURA", false);
            mEncabezadoCelda(sheet, "C", "C", 5, 0, 8, "CLIENTE", false);

            mEncabezadoCelda(sheet, "D", "D", 5, 0, 8, "MONTO TOTAL DE LA FACTURA ANTES DE IVA", false);
            mEncabezadoCelda(sheet, "E", "E", 5, 0, 8, "COSTO DE LOS PRODUCTOS", false);
            mEncabezadoCelda(sheet, "F", "F", 5, 0, 8, "UTILIDAD $$", false);
            mEncabezadoCelda(sheet, "G", "G", 5, 0, 8, "% UTILIDAD   ", false);
            mEncabezadoCelda(sheet, "H", "H", 5, 0, 8, "COMISION A PAGAR", false);
            sheet.get_Range("H" + 5.ToString(), "I" + 5.ToString()).Merge();

            sheet.get_Range("A" + 5.ToString(), "I" +
            5.ToString()).Interior.Color = Color.LightBlue;

            
            lrenglon = 6;
            //mResetearrTotales();

            string lconcepto = "";


            string lagente = "";
            //sheet.get_Range("B" + lrengloninicial, "V" + lrengloninicial).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            int lmismoconcepto = 0;
            //lrenglon += 1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos, tres;
            int lcolumna;
            string summonto = "", sumcosto="", sumutilidad="", sumcomision1="", sumcomision2="";
            string sumgmonto = "", sumgcosto = "", sumgutilidad = "", sumgcomision1 = "", sumgcomision2 = "";


            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total
                // Prog.	Fecha	Folio	Proveedor	Producto	"Cantidad Solicitada"	"Cantidad Pendiente"

                lcolumna = 1;
                if (lagente != row["CNOMBREAGENTE"].ToString().Trim())
                {
                    if (lrenglon != 6)
                    {

                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


                        sheet.Cells[lrenglon, 1].value = "TOTALES";
                        sheet.Cells[lrenglon, 4].value = summonto;
                        sheet.Cells[lrenglon, 5].value = sumcosto;
                        sheet.Cells[lrenglon, 6].value = sumutilidad;
                        sheet.Cells[lrenglon, 8].value = sumcomision1;
                        sheet.get_Range("D" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "CURRENCY";

                        sheet.get_Range("H" + lrenglon.ToString(), "I" + lrenglon.ToString()).Style = "CURRENCY";



                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                        sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;

                        sheet.get_Range("A" + (lrenglon).ToString(), "i" + lrenglon).Font.Bold = true;

                        sheet.Cells[lrenglon++, 9].value = sumcomision2;
                        

            

                    }
                    lagente = row["CNOMBREAGENTE"].ToString().Trim();

                
                    
                }
                summonto = row["summonto"].ToString().Trim();
                sumcosto = row["sumcosto"].ToString().Trim();
                sumutilidad = row["sumutilidad"].ToString().Trim();
                sumcomision1 = row["sumcomisionpropia"].ToString().Trim();
                sumcomision2 = row["sumcomision2"].ToString().Trim();

                sumgmonto = row["totalmonto"].ToString().Trim();
                sumgcosto = row["totalcosto"].ToString().Trim();
                sumgutilidad = row["totalutilidad"].ToString().Trim();
                sumgcomision1 = row["totalcomisionpropia"].ToString().Trim();
                sumgcomision2 = row["totalcomision2"].ToString().Trim();

                //NOMBRE DEL AGENTE	FACTURA	CLIENTE	 MONTO TOTAL DE LA FACTURA ANTES DE IVA 	 COSTO DE LOS PRODUCTOS 	 UTILIDAD $$ 	% UTILIDAD	 COMISION A PAGAR 	

                lcolumna = 1;
                
                //sheet.get_Range("A" + (lrenglon).ToString(), "L" + lrenglon).Font.Size = 8;
                sheet.Cells[lrenglon, lcolumna++].value = row["CNOMBREAGENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["FACTURA"].ToString().Trim();

                sheet.Cells[lrenglon, lcolumna++].value = row["CLIENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["MONTO"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["COSTO"].ToString().Trim();
               // sheet.Cells[lrenglon, lcolumna++].value = row["CCOMISIONVENTAAGENTE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["UTILIDAD"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["PORCENTAJE"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["COMISIONPROPIA"].ToString().Trim();
                sheet.Cells[lrenglon, lcolumna++].value = row["COMISION2"].ToString().Trim();

                sheet.get_Range("D" + lrenglon.ToString(), "I" + lrenglon.ToString()).Style = "CURRENCY";
                sheet.get_Range("G" + lrenglon.ToString(), "G" + lrenglon.ToString()).NumberFormat = "0.00%";



                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                sheet.get_Range("A" + (lrenglon).ToString(), "I" + (lrenglon).ToString()).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;




                //sheet.Cells[lrenglon, lcolumna++].value = "'" + fecha2; //C
                //sheet.get_Range("Q" + lrenglon.ToString(), "X" + lrenglon.ToString()).Style = "Currency";

                lrenglon++;


            }


            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


            sheet.Cells[lrenglon, 1].value = "TOTALES";
            sheet.Cells[lrenglon, 4].value = summonto;
            sheet.Cells[lrenglon, 5].value = sumcosto;
            sheet.Cells[lrenglon, 6].value = sumutilidad;
            sheet.Cells[lrenglon, 8].value = sumcomision1;
            sheet.Cells[lrenglon, 9].value = sumcomision2;
            sheet.get_Range("D" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "CURRENCY";

            sheet.get_Range("H" + lrenglon.ToString(), "I" + lrenglon.ToString()).Style = "CURRENCY";



            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;

            sheet.get_Range("A" + (lrenglon).ToString(), "i" + lrenglon).Font.Bold = true;

            //sheet.Cells[lrenglon++, 9].value = sumcomision2;


            lrenglon += 2;

            sheet.Cells[lrenglon, 1].value = "TOTALES";


            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "I" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


            sheet.Cells[lrenglon, 1].value = "TOTALES";
            sheet.Cells[lrenglon, 4].value = sumgmonto;
            sheet.Cells[lrenglon, 5].value = sumgcosto;
            sheet.Cells[lrenglon, 6].value = sumgutilidad;
            sheet.Cells[lrenglon, 8].value = sumgcomision1;
            sheet.Cells[lrenglon, 9].value = sumgcomision2;

            sheet.get_Range("D" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "CURRENCY";

            sheet.get_Range("H" + lrenglon.ToString(), "I" + lrenglon.ToString()).Style = "CURRENCY";

            sheet.get_Range("A" + (lrenglon).ToString(), "i" + lrenglon).Font.Bold = true;

            sheet.get_Range("A" + lrenglon.ToString(), "I" +
            lrenglon.ToString()).Interior.Color = Color.LightBlue;

            lrenglon += 3;
            sheet.Cells[lrenglon, 1].value = "RESUMEN COMISIONES A PAGAR";
            sheet.get_Range("A" + lrenglon.ToString(), "A" +
            lrenglon.ToString()).Interior.Color = Color.LightGreen;

            lrenglon += 3;
            sheet.Cells[lrenglon, 1].value = "AGENTES	 ";
            sheet.Cells[lrenglon, 2].value = "COMISIONES X VENTA";
            sheet.Cells[lrenglon, 3].value = "COMISION EXTRA ";
            sheet.Cells[lrenglon, 4].value = "TOTAL COMISIONES";
            sheet.get_Range("A" + (lrenglon).ToString(), "D" + lrenglon).Font.Bold = true;


            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


            lrenglon++;
            decimal totalcomision1 = 0, totalcomision2 = 0,totalcomisiones = 0;
            foreach (DataRow row in DatosResumen.Rows)
            {
                //sheet.Cells[lrenglon++, 1].value = "TOTALES";
                sheet.Cells[lrenglon, 1].value = row["cnombreagente"].ToString().Trim();
                sheet.Cells[lrenglon, 2].value = row["comision1"].ToString().Trim(); ;
                sheet.Cells[lrenglon, 3].value = row["comision2"].ToString().Trim(); ;
                sheet.get_Range("B" + lrenglon.ToString(), "D" + lrenglon.ToString()).Style = "CURRENCY";


                totalcomision1 += decimal.Parse(row["comision1"].ToString().Trim());
                totalcomision2 += decimal.Parse(row["comision2"].ToString().Trim());
                //totalcomisiones += (totalcomision1 + totalcomision2);

                sheet.Cells[lrenglon, 4].value = (decimal.Parse(row["comision1"].ToString().Trim()) + decimal.Parse(row["comision2"].ToString().Trim())).ToString();
                sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
                sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;

                lrenglon++;

                
            }

            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            sheet.get_Range("A" + lrenglon, "D" + lrenglon).Borders[MyExcel.XlBordersIndex.xlInsideVertical].LineStyle = 1;


            sheet.Cells[lrenglon, 1].value = "Totales";
            sheet.Cells[lrenglon, 2].value = totalcomision1.ToString().Trim();
            sheet.Cells[lrenglon, 3].value = totalcomision2.ToString().Trim();
            sheet.Cells[lrenglon, 4].value = (totalcomision1 + totalcomision2).ToString().Trim();

            sheet.get_Range("A" + lrenglon.ToString(), "D" +
            lrenglon.ToString()).Interior.Color = Color.LightYellow;
            sheet.get_Range("A" + (lrenglon).ToString(), "D" + lrenglon).Font.Bold = true;
            sheet.get_Range("B" + lrenglon.ToString(), "D" + lrenglon.ToString()).Style = "CURRENCY";


            sheet.Cells.EntireColumn.AutoFit();
            return;
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
    }
}
