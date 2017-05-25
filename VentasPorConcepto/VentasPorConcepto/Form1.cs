using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;


namespace VentasPorConcepto
{
    using MyExcel = Microsoft.Office.Interop.Excel; 
    public partial class Form1 : Form
    {
        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public OleDbConnection _conexion;
        long idini = 0;
        long idfin = 0;
        int lcolumnacliente = 2;
        int lcolumnaperiodo1 = 3;
        int lcolumnaperiodo2 = 4;
        int lcolumnaperiodo3 = 5;
        int lcolumnaperiodo4 = 6;
        double total1 = 0;
        double total2 = 0;
        double total3 = 0;
        double total4 = 0;
        double total5 = 0;

        

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string mensaje;
            this.comboBox1.Items.Clear();
            this.comboBox1.DataSource  = mCargarEmpresas(out mensaje);
            comboBox1.DisplayMember = "Nombre";
            comboBox1.ValueMember = "Ruta";
            try
            {
                this.comboBox1.SelectedIndex = 1;
                this.comboBox1.SelectedIndex = 0;
            }
            catch(Exception ee)
            {
                this.comboBox1.SelectedIndex = 0;
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
        public OleDbConnection mAbrirConexionOrigen()
        {
            _conexion = null;
            string rutaorigen = comboBox1.SelectedValue.ToString ();
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion.Open();
            }
            return _conexion;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            
                MyExcel.Workbook newWorkbook = mIniciarExcel();
                DataTable PorConcepto = null;
                DataTable PorEmpleado = null;
                DataTable PorClasificacion = null;
                mTraerInformacionPrimerReporte(ref PorConcepto);
                mEnviarExcelPrimerReporte(newWorkbook.Sheets[1], PorConcepto);
                //mTraerInformacionSegundoReporte(ref PorEmpleado);
                //mEnviarExcelSegundoReporte(newWorkbook.Sheets[2], PorEmpleado);
                //mTraerInformacionTercerReporte(ref PorClasificacion);
                //mEnviarExcelTercerReporte(newWorkbook.Sheets[3], PorClasificacion);
            
            

        }


        private void mPrimerReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {
            //return;
            //sheet.Cells[2, 3].value = comboBox1.Text.Trim();
            //sheet.get_Range("C2", "D2").Merge();
            //sheet.get_Range("C2").Font.Size = 16;
            //sheet.get_Range("C2").Font.Bold = true;
            //sheet.Cells[3, 3].value = "Antiguedad de Saldos por Concepto y Por Cliente";
            //sheet.get_Range("C3", "E3").Merge();
            //sheet.get_Range("C3").Font.Bold = true;

            //sheet.get_Range("B4", "G4").Merge();
            //sheet.get_Range("B4", "G4").Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;

            int lrenglon = 6;
            int lrengloninicial = 6;
            int lrengloniniciaconcepto = 6;
            int lrenglontempo = 6;

            configuracionencabezado(sheet, "Ordenes de compra", lrenglon);
            
            //mResetearrTotales();

            string lconcepto = "";


            string lcliente = "";
            sheet.get_Range("B" + lrengloninicial , "V"+ lrengloninicial ).Borders[MyExcel.XlBordersIndex.xlEdgeBottom].LineStyle = 1; 
            int lmismoconcepto = 0;
            lrenglon+=1;
            lrengloniniciaconcepto = lrenglon;
            decimal dos,tres;
            foreach (DataRow row in PorConcepto.Rows)
            {

                for (int i = 0; i < 31; i++)
                {
                    if (i==0 || i ==21 || i ==24)
                    {
                        if (row[i].ToString().Trim() != "")
                            sheet.Cells[lrenglon, i + 3].value = "'" + row[i].ToString().Trim().Substring(6, 2) + "/" + row[i].ToString().Trim().Substring(4, 2) + "/" + row[i].ToString().Trim().Substring(0, 4);
                        else
                            sheet.Cells[lrenglon, i + 3].value = "";

                      //  dos = decimal.Parse(row[i].ToString().Trim());
                       // tres = dos / decimal.Parse(row[13].ToString().Trim());
                       // sheet.Cells[lrenglon, i + 3].value = tres.ToString().Trim();
                    }
                    else
                        sheet.Cells[lrenglon, i + 3].value = row[i].ToString().Trim();
                }
                decimal totalunidades = decimal.Parse(row[31].ToString().Trim());
                decimal totalpendientes = decimal.Parse(row[32].ToString().Trim());

                if (totalpendientes == 0)
                {
                    sheet.Cells[lrenglon, 31].value = "Totalmente Surtido";
                }
                else
                {
                    if(totalpendientes == totalunidades)
                        sheet.Cells[lrenglon,31].value = "Sin surtir";
                    else
                        sheet.Cells[lrenglon, 31].value = "Parcialmente surtido";
                }
                if (row["fechavecompra"]!="")
                    sheet.Cells[lrenglon, 34].value = "'" + row["fechavecompra"].ToString().Trim().Substring(6, 2) + "/" + row["fechavecompra"].ToString().Trim().Substring(4, 2) + "/" + row["fechavecompra"].ToString().Trim().Substring(0, 4);
                //sheet.Cells[lrenglon, 34].value = row["fechavecompra"];

                // movorden.cunidades as movordenuni, " +
//" ((movorden.ctotal - movorden.cimpuesto1)/ movorden.cunidades) as costo, " +
//" 0 as costo, " +
            //" m34.cnombrem01, " +
//" movorden.cneto as netooc, movorden.cimpuesto1, movorden.ctotal as movordentotal

                if (decimal.Parse(row["movordenuni"].ToString().Trim()) ==0)
                    sheet.Cells[lrenglon, 17].value = 0;
                else
                    sheet.Cells[lrenglon, 17].value = (decimal.Parse(row["movordentotal"].ToString().Trim()) - decimal.Parse(row["movimpuesto1"].ToString().Trim())) / decimal.Parse(row["movordenuni"].ToString().Trim());
                    //"Sin surtir"; ((movorden.ctotal - movorden.cimpuesto1)/ movorden.cunidades) as costo, " +



                lrenglon++;

            }
                sheet.Cells.EntireColumn.AutoFit();
            return;
            
                
            
        }
        private void Totalizar (MyExcel.Worksheet sheet, int lrenglon, int lrengloniniciaconcepto, int lrenglontempo )
        {
            //sheet.Cells[lrenglon, 2].value = "Totales";
            sheet.get_Range("C" + lrenglon.ToString(), "C" +
            lrenglon.ToString()).Formula = "=SUM(C" + lrengloniniciaconcepto.ToString() + ":C" + lrenglontempo.ToString() + ")";
            sheet.get_Range("D" + lrenglon.ToString(), "D" +
            lrenglon.ToString()).Formula = "=SUM(D" + lrengloniniciaconcepto.ToString() + ":D" + lrenglontempo.ToString() + ")";
            sheet.get_Range("E" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Formula = "=SUM(E" + lrengloniniciaconcepto.ToString() + ":E" + lrenglontempo.ToString() + ")";
            sheet.get_Range("F" + lrenglon.ToString(), "F" +
            lrenglon.ToString()).Formula = "=SUM(F" + lrengloniniciaconcepto.ToString() + ":F" + lrenglontempo.ToString() + ")";
            sheet.get_Range("G" + lrenglon.ToString(), "G" +
            lrenglon.ToString()).Formula = "=SUM(G" + lrengloniniciaconcepto.ToString() + ":G" + lrenglontempo.ToString() + ")";

            sheet.get_Range("H" + lrenglon.ToString(), "H" +
            lrenglon.ToString()).Formula = "=SUM(G" + lrengloniniciaconcepto.ToString() + ":H" + lrenglontempo.ToString() + ")";

            sheet.get_Range("I" + lrenglon.ToString(), "I" +
            lrenglon.ToString()).Formula = "=SUM(I" + lrengloniniciaconcepto.ToString() + ":I" + lrenglontempo.ToString() + ")";

            sheet.get_Range("J" + lrenglon.ToString(), "J" +
            lrenglon.ToString()).Formula = "=SUM(J" + lrengloniniciaconcepto.ToString() + ":J" + lrenglontempo.ToString() + ")";

            sheet.get_Range("K" + lrenglon.ToString(), "K" +
            lrenglon.ToString()).Formula = "=SUM(K" + lrengloniniciaconcepto.ToString() + ":K" + lrenglontempo.ToString() + ")";

            sheet.get_Range("L" + lrenglon.ToString(), "L" +
            lrenglon.ToString()).Formula = "=SUM(L" + lrengloniniciaconcepto.ToString() + ":L" + lrenglontempo.ToString() + ")";

            sheet.get_Range("M" + lrenglon.ToString(), "M" +
            lrenglon.ToString()).Formula = "=SUM(M" + lrengloniniciaconcepto.ToString() + ":M" + lrenglontempo.ToString() + ")";
            sheet.get_Range("N" + lrenglon.ToString(), "N" +
            lrenglon.ToString()).Formula = "=SUM(N" + lrengloniniciaconcepto.ToString() + ":N" + lrenglontempo.ToString() + ")";


            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("B" + lrenglon, "G" + lrenglon).Font.Bold = true;

            total1 += double.Parse(sheet.Cells[lrenglon, 3].value.ToString());
            total2 += double.Parse(sheet.Cells[lrenglon, 4].value.ToString());
            total3 += double.Parse(sheet.Cells[lrenglon, 5].value.ToString());
            total4 += double.Parse(sheet.Cells[lrenglon, 6].value.ToString());
            total5 += double.Parse(sheet.Cells[lrenglon, 7].value.ToString());
        }
        private void configuracionencabezado(MyExcel.Worksheet sheet, string texto,int lrenglon)
        {
            sheet.Cells[2, 3].value = comboBox1.Text.Trim();
            sheet.get_Range("C2").Font.Size = 16;
            sheet.get_Range("C2").Font.Bold = true;
            sheet.get_Range("C2", "D2").Merge();
            sheet.Cells[3, 3].value = texto;
            sheet.get_Range("C3", "E3").Merge();
            sheet.get_Range("C3").Font.Bold = true;

            //Fecha (OC)	Serie	Folio	Código cliente	Razón Social	Subtotal	Impuesto 1	
            //Total	Tipo de Cambio	CLASIF1	CLASIF2	CODIGO DEL PRODUCTO	NOMBRE DEL PRODUCTO	
            //CANTIDAD ORDEN COMPRA	CANTIDAD RECIBIDA	PENDIENTE DE RECIBIR	FECHA DE RECEPCION	 
            //COSTO 	TOTAL


            sheet.Cells[lrenglon, 3].value = "Fecha (OC)";
            sheet.Cells[lrenglon, 4].value = "Serie";
            sheet.Cells[lrenglon, 5].value = "Folio";
            sheet.Cells[lrenglon, 6].value = "Codigo Cliente";
            sheet.Cells[lrenglon, 7].value = "Razon Social";
            sheet.Cells[lrenglon, 8].value = "Subtotal";
            sheet.Cells[lrenglon, 9].value = "Impuesto 1";
            sheet.Cells[lrenglon, 10].value = "Total";
            sheet.Cells[lrenglon, 11].value = "Tipo de Cambio";
            sheet.Cells[lrenglon, 12].value = "CLASIF1";
            sheet.Cells[lrenglon, 13].value = "CLASIF2";
            sheet.Cells[lrenglon, 14].value = "CODIGO DEL PRODUCTO";
            sheet.Cells[lrenglon, 15].value = "NOMBRE DEL PRODUCTO";
            sheet.Cells[lrenglon, 16].value = "CANTIDAD ORDEN DE COMPRA";

            sheet.Cells[lrenglon, 17].value = "COSTO UNITARIO";
            sheet.Cells[lrenglon, 18].value = "MONEDA O.C.";
            sheet.Cells[lrenglon, 19].value = "SUBTOTAL (MXP)";
            sheet.Cells[lrenglon, 20].value = "IVA (MXP)";
            sheet.Cells[lrenglon, 21].value = "TOTAL (MXP)";


            sheet.Cells[lrenglon, 22].value = "Serie Cons";
            sheet.Cells[lrenglon, 23].value = "Folio Cons";
            sheet.Cells[lrenglon, 24].value = "Fecha  Cons";


            sheet.Cells[lrenglon, 25].value = "Serie Compra";
            sheet.Cells[lrenglon, 26].value = "Folio Compra";
            sheet.Cells[lrenglon, 27].value = "Fecha  Compra";

            sheet.Cells[lrenglon, 28].value = "Dias";
            sheet.Cells[lrenglon, 29].value = "Cantidad solicitada";
            sheet.Cells[lrenglon, 30].value = "Cantidad Recibida";

            sheet.Cells[lrenglon, 31].value = "Status";
            sheet.Cells[lrenglon, 32].value = "Max";
            sheet.Cells[lrenglon, 33].value = "Min";

            sheet.Cells[lrenglon, 34].value = "Fecha Vencimiento";


            /*
            sheet.Cells[lrenglon, 17].value = "CANTIDAD RECIBIDA";
            sheet.Cells[lrenglon, 18].value = "PENDIENTE DE RECIBIR";
            sheet.Cells[lrenglon, 19].value = "FECHA DE RECEPCION";
            sheet.Cells[lrenglon, 20].value = "COSTO";
            sheet.Cells[lrenglon, 21].value = "TOTAL";
            */


            

            
        }
        private void mResetearrTotales()
        {
            total1 = 0;
            total2 = 0;
            total3 = 0;
            total4 = 0;
            total5 = 0;
        }

        private void mSegundoReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {

            int lrenglon = 5;
            int lrengloninicial  = 5;
            int lrengloniniciaconcepto = 5;
            int lrenglontempo = 5;
            configuracionencabezado(sheet, "Antiguedad de Saldos por Cliente y Por Concepto",lrenglon);
            mResetearrTotales();
            
            string lconcepto = "";
            string lcliente = "";
            
            
            int lmismocliente = 0;
            foreach (DataRow row in PorConcepto.Rows)
            {
                if (lcliente != row["crazonso01"].ToString().Trim())
                {
                    // sumar el total del anterior 
                    if (lrenglon != lrengloninicial)
                    {

                        sheet.get_Range("G" + lrenglon.ToString(), "G" +
                        lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";

                        sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";

                        lrenglontempo = lrenglon;
                        lrenglon++;
                        Totalizar (sheet, lrenglon, lrengloniniciaconcepto, lrenglontempo );
                    }
                    lrenglon++;
                    lcliente = row["crazonso01"].ToString().Trim();
                    sheet.Cells[lrenglon, lcolumnacliente].value = lcliente;
                    sheet.get_Range("B" + lrenglon).Font.Bold = true;
                    //sheet.get_Range("B" + lrenglon, "B" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                    
                    lrengloniniciaconcepto = lrenglon + 1;
                    lmismocliente = 0;
                }
                else
                    lmismocliente = 1;
                if (lconcepto != row["cnombrec01"].ToString().Trim())
                {
                    if (lrenglon > lrengloninicial+1 && lmismocliente == 1)
                    {
                        //    lrenglon++;
                        sheet.get_Range("G" + lrenglon.ToString(), "G" +
                        lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";
                        sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";

                    }
                    lrenglon++;
                    sheet.Cells[lrenglon, lcolumnacliente].value = row["cnombrec01"].ToString().Trim();
                    sheet.get_Range("B" + lrenglon).Font.Bold = true;
                    sheet.get_Range("B" + lrenglon, "B" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                    lconcepto = row["cnombrec01"].ToString().Trim();
                }
                else
                {
                    if (lmismocliente == 0)
                    {
                        lrenglon++;
                        sheet.Cells[lrenglon, lcolumnacliente].value = row["cnombrec01"];
                        sheet.get_Range("B" + lrenglon).Font.Bold = true;
                        sheet.get_Range("B" + lrenglon, "B" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                        lconcepto = row["cnombrec01"].ToString().Trim ();
                    }
                }
                
                switch (row["Orden"].ToString())
                {
                    case "1":
                        sheet.Cells[lrenglon, lcolumnaperiodo1].value = row["sum_cpendiente"];
                        break;
                    case "2":
                        sheet.Cells[lrenglon, lcolumnaperiodo2].value = row["sum_cpendiente"];
                        break;
                    case "3":
                        sheet.Cells[lrenglon, lcolumnaperiodo3].value = row["sum_cpendiente"];
                        break;
                    case "4":
                        sheet.Cells[lrenglon, lcolumnaperiodo4].value = row["sum_cpendiente"];
                        break;
                }

            }
            sheet.get_Range("G" + lrenglon.ToString(), "G" +
                            lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";

            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("B" + lrenglon, "B" + lrenglon).Font.Bold = true;

            lrenglontempo = lrenglon;
            lrenglon++;
            Totalizar(sheet, lrenglon, lrengloniniciaconcepto, lrenglontempo);
            
            lrenglon++;
            sheet.Cells[lrenglon, 2].value = "Totales Generales";
            sheet.Cells[lrenglon, 3].value = total1;
            sheet.Cells[lrenglon, 4].value = total2;
            sheet.Cells[lrenglon, 5].value = total3;
            sheet.Cells[lrenglon, 6].value = total4;
            sheet.Cells[lrenglon, 7].value = total5;

            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Font.Bold = true;
            sheet.Cells.EntireColumn.AutoFit();
        }

        private void mTercerReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {
            int lrenglon = 5;
            int lrengloniniciaconcepto = 5;
            int lrenglontempo = 5;
            int lrengloinicial = 5;

            configuracionencabezado(sheet, "Antiguedad de Saldos por Clasificacion y Por Cliente", lrenglon);
            mResetearrTotales();
            

            int lcolumnacliente = 2;
            int lcolumnaperiodo1 = 3;
            int lcolumnaperiodo2 = 4;
            int lcolumnaperiodo3 = 5;
            int lcolumnaperiodo4 = 6;

            
            string lconcepto = "";
            string lcliente = "";
            
            int lmismoconcepto = 0;
            double total1 = 0;
            double total2 = 0;
            double total3 = 0;
            double total4 = 0;
            double total5 = 0;


            foreach (DataRow row in PorConcepto.Rows)
            {
                if (lconcepto != row["cvalorcl01"].ToString().Trim())
                {
                    if (lrenglon > lrengloinicial )
                    {
                        if (lrenglon > lrengloinicial + 2)
                        {
                            sheet.get_Range("G" + lrenglon.ToString(), "G" +
                            lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";
                            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
                        }


                        sheet.Cells[lrenglon, lcolumnacliente].value = lcliente;
                        lcliente = row["crazonso01"].ToString().Trim ();

                        lrenglontempo = lrenglon;
                        lrenglon++;
                        Totalizar(sheet, lrenglon, lrengloniniciaconcepto, lrenglontempo);
                    }
                    lrenglon++;
                    sheet.Cells[lrenglon, lcolumnacliente].value = row["cvalorcl01"];
                    sheet.get_Range("B" + lrenglon, "B" + lrenglon).Font.Bold = true;
                    sheet.get_Range("B" + lrenglon, "B" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
                    lconcepto = row["cvalorcl01"].ToString().Trim();
                    

                    lrengloniniciaconcepto = lrenglon + 1;
                    lmismoconcepto = 0;
                }
                else
                    lmismoconcepto = 1;
                if (lcliente != row["crazonso01"].ToString().Trim())
                {
                    // sumar el total del anterior 
                    if (lrenglon > 6)
                    {
                        sheet.get_Range("G" + lrenglon.ToString(), "G" +
                        lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";
                        sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
                    }
                    lrenglon++;
                    sheet.Cells[lrenglon, lcolumnacliente].value = row["crazonso01"].ToString().Trim ();
                    lcliente = row["crazonso01"].ToString();

                }
                else
                {
                    if (lmismoconcepto == 0)
                    {
                        lrenglon++;
                        sheet.Cells[lrenglon, lcolumnacliente].value = row["crazonso01"].ToString().Trim();
                        lcliente = row["crazonso01"].ToString().Trim();
                    }
                }
                switch (row["Orden"].ToString())
                {
                    case "1":
                        sheet.Cells[lrenglon, lcolumnaperiodo1].value = row["sum_cpendiente"];
                        break;
                    case "2":
                        sheet.Cells[lrenglon, lcolumnaperiodo2].value = row["sum_cpendiente"];
                        break;
                    case "3":
                        sheet.Cells[lrenglon, lcolumnaperiodo3].value = row["sum_cpendiente"];
                        break;
                    case "4":
                        sheet.Cells[lrenglon, lcolumnaperiodo4].value = row["sum_cpendiente"];
                        break;
                }

            }
            sheet.get_Range("G" + lrenglon.ToString(), "G" +
                            lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";
            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
            lrenglontempo = lrenglon;
            lrenglon++;
            Totalizar(sheet, lrenglon, lrengloniniciaconcepto, lrenglontempo);

            lrenglon++;
            sheet.Cells[lrenglon, 2].value = "GRAN TOTAL";
            sheet.Cells[lrenglon, 3].value = total1 ;
            sheet.Cells[lrenglon, 4].value = total2;
            sheet.Cells[lrenglon, 5].value = total3;
            sheet.Cells[lrenglon, 6].value = total4;
            sheet.Cells[lrenglon, 7].value = total5;
            sheet.get_Range("C" + lrenglon.ToString(), "G" + lrenglon.ToString()).Style = "Currency";
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Borders[MyExcel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            sheet.get_Range("C" + lrenglon, "G" + lrenglon).Font.Bold = true;

            sheet.Cells.EntireColumn.AutoFit();
            
        }

        private void ltotalizareporte3(int lrenglon, int lultimosub, MyExcel.Worksheet sheet, int lrengloniniciaconcepto)
        {
            if (lultimosub ==1)
                sheet.get_Range("G" + lrenglon.ToString(), "G" +
                            lrenglon.ToString()).Formula = "=SUM(C" + lrenglon.ToString() + ":F" + lrenglon.ToString() + ")";

            int lrenglontempo = lrenglon;
            lrenglon++;
            sheet.Cells[lrenglon, 2].value = "Totales";
            sheet.get_Range("C" + lrenglon.ToString(), "C" +
            lrenglon.ToString()).Formula = "=SUM(C" + lrengloniniciaconcepto.ToString() + ":C" + lrenglontempo.ToString() + ")";
            sheet.get_Range("D" + lrenglon.ToString(), "D" +
            lrenglon.ToString()).Formula = "=SUM(D" + lrengloniniciaconcepto.ToString() + ":D" + lrenglontempo.ToString() + ")";
            sheet.get_Range("E" + lrenglon.ToString(), "E" +
            lrenglon.ToString()).Formula = "=SUM(E" + lrengloniniciaconcepto.ToString() + ":E" + lrenglontempo.ToString() + ")";
            sheet.get_Range("F" + lrenglon.ToString(), "F" +
            lrenglon.ToString()).Formula = "=SUM(F" + lrengloniniciaconcepto.ToString() + ":F" + lrenglontempo.ToString() + ")";
            sheet.get_Range("G" + lrenglon.ToString(), "G" +
            lrenglon.ToString()).Formula = "=SUM(G" + lrengloniniciaconcepto.ToString() + ":G" + lrenglontempo.ToString() + ")";
        }



        private void mEnviarExcelPrimerReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {
            mPrimerReporte(sheet, PorConcepto);
        }
        private void mEnviarExcelSegundoReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {
            mSegundoReporte(sheet, PorConcepto);
        }

        private void mEnviarExcelTercerReporte(MyExcel.Worksheet sheet, DataTable PorConcepto)
        {
            mTercerReporte(sheet, PorConcepto);
        }


        private void mTraerInformacionPrimerReporte(ref DataTable aDS)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen();

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() +  lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            string lquery = "select top 100 substr(dtos(m8.cfecha),7,2) + '/' + substr(dtos(m8.cfecha),5,2) + '/' + substr(dtos(m8.cfecha),1,4), cseriedo01, cfolio, m2.ccodigoc01, m2.crazonso01, m8.ctotal - m8.cimpuesto1, " +
" m8.cimpuesto1, m8.ctotal,m8.ctipocam01, m20.cvalorcl01, m20a.cvalorcl01,m5.ccodigop01, m5.cnombrep01, m10.cunidades, " +
                //" m10.cunidades -m10.cunidade03, m10a.cunidade05, substr(dtos(m10a.cfecha),7,2) + '/' + substr(dtos(m10a.cfecha),5,2) + '/' + substr(dtos(m10a.cfecha),1,4)   " +
"nvl(m10a.cunidade05,0),m10.cunidades-nvl(m10a.cunidade05,0) ,  substr(dtos(m10a.cfecha),7,2) + '/' + substr(dtos(m10a.cfecha),5,2) + '/' + substr(dtos(m10a.cfecha),1,4) " +
" , m10.ctotal - m10.cimpuesto1, m10.ctotal" +
        " from mgw10008 m8 " +
" join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
" join mgw10010 m10 on m8.ciddocum01 = m10.ciddocum01 " +
" left join mgw10010 m10a on m10.cidmovim01 = m10a.cidmovto02 and m10a.ciddocum02 = 18" +
" join mgw10005 m5 on m5.cidprodu01 = m10.cidprodu01 " +
" join mgw10020 m20 on m20.cidvalor01 = m5.cidvalor01 " +
" join mgw10020 m20a on m20a.cidvalor01 = m5.cidvalor02 " +
" where m8.ciddocum02 = 17 " +
" and dtos(m8.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' order by m8.cfecha, m8.cfolio";;


            

            
             //string lquery1 = "select substr(dtos(oc.cfecha),7,2) + '/' + substr(dtos(oc.cfecha),5,2) + '/' + substr(dtos(oc.cfecha),1,4)," +
            string lquery1 = "select dtos(oc.cfecha)," +
" oc.cseriedo01 as serieoc, oc.cfolio as foliooc, m2.ccodigoc01, m2.crazonso01, oc.ctotal - oc.cimpuesto1 as netooc,  oc.cimpuesto1, oc.ctotal,oc.ctipocam01,  " +
" m20.cvalorcl01, m20a.cvalorcl01,m5.ccodigop01, m5.cnombrep01,   movorden.cunidades as movordenuni, " +
//" ((movorden.ctotal - movorden.cimpuesto1)/ movorden.cunidades) as costo, " +
" 0 as costo, " +
            " m34.cnombrem01, " +
" movorden.cneto as netooc, movorden.cimpuesto1 as movimpuesto1, movorden.ctotal as movordentotal,cons.cseriedo01 as seriecons, cons.cfolio as foliocons, " +
            //" substr(dtos(cons.cfecha),7,2) + '/' + substr(dtos(cons.cfecha),5,2) + '/' + substr(dtos(cons.cfecha),1,4) as fechacons, compra.cseriedo01 as seriecompra, compra.cfolio as foliocompra, " +
            " dtos(cons.cfecha) as fechacons, compra.cseriedo01 as seriecompra, compra.cfolio as foliocompra, " +
            //" substr(dtos(compra.cfecha),7,2) + '/' + substr(dtos(compra.cfecha),5,2) + '/' + substr(dtos(compra.cfecha),1,4) as fechacompra, 0 as dias, " +
            " dtos(compra.cfecha) as fechacompra, 0 as dias, " +
" movorden.cunidades, movcompra.cunidades, '' as status, m16.cexistmi01, m16.cexistma01,oc.ctotalun01, oc.cunidade01  " +
//" ,substr(dtos(oc.cfechave01),7,2) + '/' + substr(dtos(oc.cfechave01),5,2) + '/' + substr(dtos(oc.cfechave01),1,4) as fechavecompra" +
" ,dtos(oc.cfechave01) as fechavecompra" +
" from mgw10008 oc  join mgw10002 m2 on oc.cidclien01 = m2.cidclien01   " +
" join mgw10034 m34 on oc.cidmoneda = m34.cidmoneda  " +
" join mgw10010 movorden on oc.ciddocum01 = movorden.ciddocum01   " +
" join mgw10005 m5 on m5.cidprodu01 = movorden.cidprodu01   " +
" left join mgw10010 movcons on movorden.cidmovim01 = movcons.cidmovto02   " +
" left join mgw10008 cons on cons.ciddocum01 = movcons.ciddocum01  " +
" left join mgw10010 movcompra on movcons.cidmovim01 = movcompra.cidmovto02   " +
" left join mgw10008 compra on compra.ciddocum01 = movcompra.ciddocum01  " +
" left join mgw10016 m16 on m16.cidprodu01 = movorden.cidprodu01 and m16.cidalmacen =1  " +
" join mgw10020 m20 on m20.cidvalor01 = m5.cidvalor01  " +
" join mgw10020 m20a on m20a.cidvalor01 = m5.cidvalor02  " +
" where oc.ciddocum02 = 17 "  +
" and dtos(oc.cfecha) between '" + sfecha1 + "' and '" + sfecha2 + "' order by oc.cfecha, oc.cfolio";

            //and month(m8.cfecha) = 6 order by m8.cfecha, m8.cfolio";

            //lquery += ejercicio;
            //lquery += " join mgw10006 m6 on m6.cidconce01 = m8.cidconce01 ";
            //lquery += sconceptos;
            //lquery += " group by 4,3,6 order by m6.cnombrec01 ";
            OleDbCommand mySqlCommand = new OleDbCommand(lquery1, lconexion );

            
            DataSet ds = new DataSet();

            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            aDS = ds.Tables[0];

        }

        private void mTraerInformacionSegundoReporte(ref DataTable aDS)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen();

            DateTime lfecha = dateTimePicker1.Value;
            double lcuantosdias = 0; // double.Parse(textBox5.Text) * -1;

            DateTime lfechaperiodo1 = lfecha;
            DateTime lfechaperiodo2 = lfechaperiodo1.AddDays(lcuantosdias);
            DateTime lfechaperiodo3 = lfechaperiodo2.AddDays(-1);
            DateTime lfechaperiodo4 = lfechaperiodo3.AddDays(lcuantosdias);
            DateTime lfechaperiodo5 = lfechaperiodo4.AddDays(-1);
            DateTime lfechaperiodo6 = lfechaperiodo5.AddDays(lcuantosdias);
            DateTime lfechaperiodo7 = lfechaperiodo6.AddDays(-1);
            DateTime lfechaperiodo8 = lfechaperiodo7.AddDays(lcuantosdias);

            string sfechaperiodo1 = lfechaperiodo1.ToString().Substring(6, 4) + lfechaperiodo1.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo1.ToString().Substring(0, 2);
            string sfechaperiodo2 = lfechaperiodo2.ToString().Substring(6, 4) + lfechaperiodo2.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo2.ToString().Substring(0, 2);
            string sfechaperiodo3 = lfechaperiodo3.ToString().Substring(6, 4) + lfechaperiodo3.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo3.ToString().Substring(0, 2);
            string sfechaperiodo4 = lfechaperiodo4.ToString().Substring(6, 4) + lfechaperiodo4.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo4.ToString().Substring(0, 2);
            string sfechaperiodo5 = lfechaperiodo5.ToString().Substring(6, 4) + lfechaperiodo5.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo5.ToString().Substring(0, 2);
            string sfechaperiodo6 = lfechaperiodo6.ToString().Substring(6, 4) + lfechaperiodo6.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo6.ToString().Substring(0, 2);
            string sfechaperiodo7 = lfechaperiodo7.ToString().Substring(6, 4) + lfechaperiodo7.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo7.ToString().Substring(0, 2);
            string sfechaperiodo8 = lfechaperiodo8.ToString().Substring(6, 4) + lfechaperiodo8.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo8.ToString().Substring(0, 2);

            string sconceptos = " and m8.cidconce01 in (";
            ListBox.SelectedObjectCollection lista;
            
            if (sconceptos == " and m8.cidconce01 in (")
                sconceptos = "";
            else
            {
                sconceptos = sconceptos.Substring(0, sconceptos.Length - 1);
                sconceptos += ")";
            }







            string lquery = "select m2.crazonso01, sum(cpendiente), m6.cnombrec01, '1' as orden " +
" from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01" +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
" where cpendiente > 0" +
" and ccancelado = 0" +
" and dtos(m8.cfecha) >= '" + sfechaperiodo2 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo1 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m6.cnombrec01, m2.crazonso01" +
" union" +
" select m2.crazonso01, sum(cpendiente), m6.cnombrec01, '2' as orden" +
" from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01" +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
" where cpendiente > 0" +
" and ccancelado = 0" +
" and dtos(m8.cfecha) >= '" + sfechaperiodo4 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo3 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m6.cnombrec01, m2.crazonso01" +
" union" +
" select m2.crazonso01, sum(cpendiente), m6.cnombrec01, '3' as orden" +
" from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01" +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
" where cpendiente > 0" +
" and ccancelado = 0" +
" and dtos(m8.cfecha) >= '" + sfechaperiodo6 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo5 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m6.cnombrec01, m2.crazonso01" +
" union" +
" select m2.crazonso01, sum(cpendiente), m6.cnombrec01, '4' as orden" +
" from mgw10008 m8 join mgw10006 m6 on m8.cidconce01 = m6.cidconce01" +
" join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " +
" where cpendiente > 0" +
" and ccancelado = 0" +
" and dtos(m8.cfecha) <= '" + sfechaperiodo7 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m6.cnombrec01, m2.crazonso01" +
" order by 1,3,4 ";

            OleDbCommand mySqlCommand = new OleDbCommand(lquery, lconexion);

            //SqlParameter month = new SqlParameter();
            //SqlParameter year = new SqlParameter();

            /*month.Value = comboBox1.SelectedIndex + 1;
            month.ParameterName = "@month";
            year.Value = comboBox2.SelectedItem;
            year.ParameterName = "@year";*/

            DataSet ds = new DataSet();

            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            //mySqlCommand.Parameters.Add(year);
            //mySqlCommand.Parameters.Add(month);
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            aDS = ds.Tables[0];

        }


        private void mTraerInformacionTercerReporte(ref DataTable aDS)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen();

            DateTime lfecha = dateTimePicker1.Value;
            double lcuantosdias = 0; // double.Parse(textBox5.Text) * -1;

            DateTime lfechaperiodo1 = lfecha;
            DateTime lfechaperiodo2 = lfechaperiodo1.AddDays(lcuantosdias);
            DateTime lfechaperiodo3 = lfechaperiodo2.AddDays(-1);
            DateTime lfechaperiodo4 = lfechaperiodo3.AddDays(lcuantosdias);
            DateTime lfechaperiodo5 = lfechaperiodo4.AddDays(-1);
            DateTime lfechaperiodo6 = lfechaperiodo5.AddDays(lcuantosdias);
            DateTime lfechaperiodo7 = lfechaperiodo6.AddDays(-1);
            DateTime lfechaperiodo8 = lfechaperiodo7.AddDays(lcuantosdias);

            string sfechaperiodo1 = lfechaperiodo1.ToString().Substring(6, 4) + lfechaperiodo1.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo1.ToString().Substring(0, 2);
            string sfechaperiodo2 = lfechaperiodo2.ToString().Substring(6, 4) + lfechaperiodo2.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo2.ToString().Substring(0, 2);
            string sfechaperiodo3 = lfechaperiodo3.ToString().Substring(6, 4) + lfechaperiodo3.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo3.ToString().Substring(0, 2);
            string sfechaperiodo4 = lfechaperiodo4.ToString().Substring(6, 4) + lfechaperiodo4.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo4.ToString().Substring(0, 2);
            string sfechaperiodo5 = lfechaperiodo5.ToString().Substring(6, 4) + lfechaperiodo5.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo5.ToString().Substring(0, 2);
            string sfechaperiodo6 = lfechaperiodo6.ToString().Substring(6, 4) + lfechaperiodo6.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo6.ToString().Substring(0, 2);
            string sfechaperiodo7 = lfechaperiodo7.ToString().Substring(6, 4) + lfechaperiodo7.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo7.ToString().Substring(0, 2);
            string sfechaperiodo8 = lfechaperiodo8.ToString().Substring(6, 4) + lfechaperiodo8.ToString().Substring(3, 2).PadLeft(2, '0') + lfechaperiodo8.ToString().Substring(0, 2);

            string sconceptos = " and m8.cidconce01 in (";
            ListBox.SelectedObjectCollection lista;
            
            if (sconceptos == " and m8.cidconce01 in (")
                sconceptos = "";
            else
            {
                sconceptos = sconceptos.Substring(0, sconceptos.Length - 1);
                sconceptos += ")";
            }


            string lquery = "select m2.crazonso01, sum(cpendiente), m20.cvalorcl01, '1' as orden " + 
" from mgw10008 m8 join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " + 
" join mgw10020 m20 on m20.cidvalor01 = m2.cidvalor01 " + 
" and m20.cidclasi01 = 7 " + 
" where cpendiente > 0 " + 
" and ccancelado = 0 " +
" and dtos(m8.cfecha) >= '" + sfechaperiodo2 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo1 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m20.cvalorcl01, m2.crazonso01 " + 
" union " + 
" select m2.crazonso01, sum(cpendiente), m20.cvalorcl01, '2' as orden " + 
" from mgw10008 m8 join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " + 
" join mgw10020 m20 on m20.cidvalor01 = m2.cidvalor01 " + 
" and m20.cidclasi01 = 7 " + 
" where cpendiente > 0 " + 
" and ccancelado = 0 " +
" and dtos(m8.cfecha) >= '" + sfechaperiodo4 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo3 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m20.cvalorcl01, m2.crazonso01 " + 
" union  " + 
" select m2.crazonso01, sum(cpendiente), m20.cvalorcl01, '3' as orden  " + 
" from mgw10008 m8 join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " + 
" join mgw10020 m20 on m20.cidvalor01 = m2.cidvalor01 " + 
" and m20.cidclasi01 = 7 " + 
" where cpendiente > 0 " + 
" and ccancelado = 0 " +
" and dtos(m8.cfecha) >= '" + sfechaperiodo6 + "' and dtos(m8.cfecha) <= '" + sfechaperiodo5 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m20.cvalorcl01, m2.crazonso01 " + 
" union  " + 
" select m2.crazonso01, sum(cpendiente), m20.cvalorcl01, '4' as orden " + 
" from mgw10008 m8 join mgw10002 m2 on m2.cidclien01 = m8.cidclien01 " + 
" join mgw10020 m20 on m20.cidvalor01 = m2.cidvalor01 " + 
" and m20.cidclasi01 = 7 " + 
" where cpendiente > 0 " + 
" and ccancelado = 0 " +
" and dtos(m8.cfecha) <='" + sfechaperiodo7 + "'";
            lquery += sconceptos;
            if (idini > 0 && idfin > 0)
            {
                lquery += " and m2.cidclien01 >= " + idini + " and m2.cidclien01 <= " + idfin;
            }
            lquery += " group by m20.cvalorcl01, m2.crazonso01 " + 
" order by 3,1,4 ";

            OleDbCommand mySqlCommand = new OleDbCommand(lquery, lconexion);

            //SqlParameter month = new SqlParameter();
            //SqlParameter year = new SqlParameter();

            /*month.Value = comboBox1.SelectedIndex + 1;
            month.ParameterName = "@month";
            year.Value = comboBox2.SelectedItem;
            year.ParameterName = "@year";*/

            DataSet ds = new DataSet();

            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            //mySqlCommand.Parameters.Add(year);
            //mySqlCommand.Parameters.Add(month);
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            aDS = ds.Tables[0];

        }

        private MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            return newWorkbook;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public List<RegConcepto> mCargarConceptos()
        {
            OleDbConnection lconexion = new OleDbConnection();

            lconexion = mAbrirConexionOrigen();
            List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                OleDbCommand lsql = new OleDbCommand("select cidconce01,ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = 4", lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegFacturas.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegConcepto lRegConcepto = new RegConcepto();
                        lRegConcepto.Codigo = lreader[1].ToString();
                        lRegConcepto.Nombre = lreader[2].ToString();
                        lRegConcepto.id =  long.Parse ( lreader[0].ToString());
                        _RegFacturas.Add(lRegConcepto);
                    }
                }

                lreader.Close();
                
            }

            
            return _RegFacturas;



        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        

        public void mBuscarCliente(string aCodigo, ref long id, ref string crazonsocial)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string lcadena="";
            id = 0;
            crazonsocial = "";

            lconexion = mAbrirConexionOrigen();
            if (lconexion != null)
            {
                // este es para flexo
                OleDbCommand lsql = new OleDbCommand("select cidclien01, crazonso01 from mgw10002 where ccodigoc01 = '" + aCodigo + "'", lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    id = long.Parse( lreader[0].ToString());
                    crazonsocial = lreader[1].ToString();
                }
                lreader.Close();

            }
            



        }


    }
}
