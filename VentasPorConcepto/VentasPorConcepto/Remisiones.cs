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

namespace VentasPorConcepto
{
    public partial class Remisiones : Form
    {

        protected ClassRN lrn = new ClassRN();
        public string Cadenaconexion = "";
        public string Archivo = "";
         Class1 x = new Class1();


        private void OnComboChange(object sender, EventArgs e)
        {
            //MessageBox.Show("cambia");
            //lrn.almacenes = "chido";
            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            //Properties.Settings.Default.database = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();
            codigocatalogocomercial1.lrn.lbd.cadenaconexion = Cadenaconexion;
            codigocatalogocomercial1.lrn = lrn;

        }

        public Remisiones()
        {
            InitializeComponent();
            empresasComercial1.SelectedItem += new EventHandler(OnComboChange);
            codigocatalogocomercial1.TextLeave += new EventHandler(OnTextLeave1);
            codigocatalogocomercial2.TextLeave += new EventHandler(OnTextLeave2);
            codigocatalogocomercial3.TextLeave += new EventHandler(OnTextLeave3);
            codigocatalogocomercial4.TextLeave += new EventHandler(OnTextLeave4);
            codigocatalogocomercial1.mSetLabelText("Cliente Inicial");
            codigocatalogocomercial2.mSetLabelText("Cliente Final");
            codigocatalogocomercial3.mSetLabelText("Agente Inicial");
            codigocatalogocomercial4.mSetLabelText("Agente Final");


            codigocatalogocomercial5.TextLeave += new EventHandler(OnTextLeave5);
            codigocatalogocomercial6.TextLeave += new EventHandler(OnTextLeave6);

            codigocatalogocomercial5.mSetLabelText("Agente Inicial");
            codigocatalogocomercial6.mSetLabelText("Agente Final");


            codigocatalogocomercial7.TextLeave += new EventHandler(OnTextLeave7);
            codigocatalogocomercial7.mSetLabelText("Serie");

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
            if (codigocatalogocomercial2.mGetCodigo()!="")
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
                RegCliente lcliente = x.mValidarCatalogoComercial(2, codigocatalogocomercial3.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial3.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Agente no Existe");
                    codigocatalogocomercial3.mSetFocus();
                }
            }
        }
        public void OnTextLeave4(object sender, EventArgs e)
        {
            if (codigocatalogocomercial4.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(2, codigocatalogocomercial4.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial4.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Agente no Existe");
                    codigocatalogocomercial4.mSetFocus();
                }
            }
        }

        public void OnTextLeave5(object sender, EventArgs e)
        {
            if (codigocatalogocomercial5.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(2, codigocatalogocomercial5.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial5.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Agente no Existe");
                    codigocatalogocomercial5.mSetFocus();
                }
            }
        }

        public void OnTextLeave6(object sender, EventArgs e)
        {
            if (codigocatalogocomercial6.mGetCodigo() != "")
            {
                RegCliente lcliente = x.mValidarCatalogoComercial(2, codigocatalogocomercial6.mGetCodigo(), empresasComercial1.aliasbdd);
                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial6.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Agente no Existe");
                    codigocatalogocomercial6.mSetFocus();
                }
            }
        }


        public void OnTextLeave7(object sender, EventArgs e)
        {
           /* if (codigocatalogocomercial7.mGetCodigo() != ""){
           
                //RegCliente lcliente = x.mValidarCatalogoComercial(2, codigocatalogocomercial7.mGetCodigo(), empresasComercial1.aliasbdd);

                RegCliente lcliente = x.mValidarCatalogoComercial(6, codigocatalogocomercial7.mGetCodigo(), empresasComercial1.aliasbdd);

                if (lcliente.RazonSocial != "")
                    codigocatalogocomercial7.mSetDescripcion(lcliente.RazonSocial);
                else
                {
                    MessageBox.Show("Usuario no Existe");
                    //codigocatalogocomercial7.mSetDescripcion= "";
                    codigocatalogocomercial7.mSetFocus();
                }
            }*/
        }




        private void Remisiones_Load(object sender, EventArgs e)
        {

<<<<<<< HEAD
            this.Text = " Reportes Ferquiagro " + " " + this.ProductVersion;
=======
            this.Text = " Reporte/Borrado Remisiones " + " " + this.ProductVersion;
>>>>>>> 9a41ea45bd8e9002eb6a577c27983ff67c519b3f
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());


            //this.codigocatalogocomercial1.mSetLibreria(lrn, Cadenaconexion);

           // this.codigocatalogocomercial1.mSetLibreria(lrn);
           // this.codigocatalogocomercial1.mSeteartipo(1, 1);

            


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
            {
                empresasComercial1.Populate(Cadenaconexion);
                RegConexion x = new RegConexion();
                x.database = Properties.Settings.Default.database;
                x.server = Properties.Settings.Default.server;
                x.usuario = Properties.Settings.Default.user;
                x.ps= Properties.Settings.Default.password;
            //    lrn.lbd.miconexion._


                //this.codigocatalogocomercial1.mSetConexion(x);
            }
            else
            {
                Form4 x = new Form4();
                x.Show();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            // string lquery;



            StringBuilder lquery = new StringBuilder();

            lquery.Append("SELECT format(d.CFECHA,'dd/MM/yyyy') as FECHA, d.cfolio as FOLIO, c.CRAZONSOCIAL AS [RAZON SOCIAL],d.CTOTALUNIDADES AS [TOTAL UNIDADES], d.cneto AS NETO, d.CTOTAL AS TOTAL,co.ccodigoconcepto, d.cseriedocumento as SERIE ");
            lquery.Append("FROM admDocumentos d ");
            lquery.Append("JOIN admClientes c on c.CIDCLIENTEPROVEEDOR = d.CIDCLIENTEPROVEEDOR ");
            lquery.Append("JOIN admConceptos co on co.CIDCONCEPTODOCUMENTO = d.CIDCONCEPTODOCUMENTO ");
            lquery.Append("where d.CCANCELADO = 0 and d.CIDDOCUMENTODE = 3 ");
            lquery.Append("and d.CFECHA between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append("order by Fecha ");

            lquery.Append("SELECT  p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, SUM(m.CUNIDADES) AS CUNIDADES, SUM(m.CNETO) AS CNETO, SUM(m.ctotal) AS CTOTAL ");
            lquery.Append("FROM admDocumentos d ");
            lquery.Append("JOIN admMovimientos m on d.CIDDOCUMENTO = m.CIDDOCUMENTO ");
            lquery.Append("join admProductos p on p.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("where d.CCANCELADO = 0 and d.CIDDOCUMENTODE = 3 ");
            lquery.Append("and d.CFECHA between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append(" group by p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO ");
            lquery.Append("order by  p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO ;");
            
            

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);

            dataGridView1.DataSource = null;
            dataGridView1.DataSource = x.DatosReporte;
            dataGridView1.AutoResizeColumns();
            //dataGridView1.AutoGenerateColumns = false;
            /*DataGridViewCheckBoxColumn CBColumn = new DataGridViewCheckBoxColumn();
            CBColumn.HeaderText = "ColumnHeader";
            CBColumn.FalseValue = "0";
            CBColumn.TrueValue = "1";
            dataGridView1.Columns.Insert(0, CBColumn);*/

            dataGridView1.Columns[6].Visible = false;

            
            if (checkBox1.Checked == true)
                x.mReporteRemisionesComercial(empresasComercial1.aliasbdd, dateTimePicker1.Value, dateTimePicker2.Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            

            Properties.Settings.Default.RutaEmpresaADM = empresasComercial1.aliasbdd;
            Properties.Settings.Default.Save();

            //lr.mborr
            RegConexion newcon = new RegConexion();
            newcon.database = empresasComercial1.aliasbdd;
            newcon.server = Properties.Settings.Default.server;
            newcon.usuario = Properties.Settings.Default.user;
            newcon.ps = Properties.Settings.Default.password;
            lrn.mAsignaEmpresaComercial(newcon);
            foreach (DataGridViewRow x in dataGridView1.Rows)
            {
                
                lrn.mBorrarDocto(x.Cells["ccodigoconcepto"].Value.ToString(), x.Cells["SERIE"].Value.ToString(), x.Cells["FOLIO"].Value.ToString());
            }
            MessageBox.Show("Proceso Terminado");
        }

        private void Remisiones_FormClosed(object sender, FormClosedEventArgs e)
        {
            lrn.mCerrarSdkComercial();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker3.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');


            DateTime lfecha2 = dateTimePicker4.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');
            
            // string lquery;



            StringBuilder lquery = new StringBuilder();

            lquery.Append("select  c.CCODIGOCLIENTE, c.CRAZONSOCIAL, c.CDIASCREDITOCLIENTE, a.CNOMBREAGENTE, d.CSERIEDOCUMENTO, d.CFOLIO ");
            lquery.Append(",d.cfecha as cfechad, p.cfecha as cfechap, d.CTOTAL, d.CPENDIENTE , isnull(datediff(day,d.cfecha, p.cfecha),-1) as numdias ");
            lquery.Append(", m.CUNIDADESCAPTURADAS, u.CABREVIATURA, pr.CCODIGOPRODUCTO, pr.CNOMBREPRODUCTO, m.ctotal as ctotalmov ");
            lquery.Append(",a.CTIPOAGENTE ");
            lquery.Append(", case  ");
            lquery.Append("when clag.ccodigovalorclasificacion  = '1' then isnull(pr.CIMPORTEEXTRA2,0) ");
            lquery.Append("when clag.ccodigovalorclasificacion  = '2' then isnull(pr.CIMPORTEEXTRA3,0) ");
            lquery.Append("when clag.ccodigovalorclasificacion  = '3' then isnull(pr.CIMPORTEEXTRA4,0) ");
            lquery.Append("else 0 ");
            lquery.Append("end as comision  ");
            lquery.Append(", cl.CVALORCLASIFICACION ");
            lquery.Append(" from admdocumentos d ");
            lquery.Append("join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
            lquery.Append("join admAgentes a on a.CIDAGENTE = d.CIDAGENTE ");
            lquery.Append("left join admAsocCargosAbonos ca on ca.CIDDOCUMENTOCARGO = d.CIDDOCUMENTO ");
            lquery.Append("left join admDocumentos p on ca.CIDDOCUMENTOABONO = p.CIDDOCUMENTO ");
            lquery.Append("join admMovimientos m on m.CIDDOCUMENTO = d.CIDDOCUMENTO ");
            lquery.Append("join admProductos pr on pr.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("join admUnidadesMedidaPeso u on m.CIDUNIDAD = u.CIDUNIDAD ");
            lquery.Append("join admClasificacionesValores cl on cl.CIDVALORCLASIFICACION = pr.CIDVALORCLASIFICACION1 ");
            lquery.Append("join admClasificacionesValores clag on clag.CIDVALORCLASIFICACION = a.CIDVALORCLASIFICACION1 ");
            lquery.Append("where d.CIDDOCUMENTODE = 4 ");
            lquery.Append("and d.ccancelado = 0 ");


            lquery.Append("and d.CFECHA >= '" + sfecha1 + "'");
            lquery.Append("and d.CFECHA <= '" + sfecha2 + "'");

            if (codigocatalogocomercial1.mGetCodigo() != "" && codigocatalogocomercial2.mGetCodigo() != "")
            {
                lquery.Append("and c.ccodigocliente >= '" + codigocatalogocomercial1.mGetCodigo() + "'");
                lquery.Append("and c.ccodigocliente <= '" + codigocatalogocomercial2.mGetCodigo() + "'");

            }

            if (codigocatalogocomercial3.mGetCodigo() != "" && codigocatalogocomercial4.mGetCodigo() != "")
            {
                lquery.Append("and a.ccodigoagente >= '" + codigocatalogocomercial3.mGetCodigo() + "'");
                lquery.Append("and a.ccodigoagente <= '" + codigocatalogocomercial4.mGetCodigo() + "'");

            }
            /*lquery.Append("and d.CFECHA between '" + sfecha1 + "' and '" + sfecha2 + "' ");
            lquery.Append("order by Fecha ");
            */
            lquery.Append("order by d.cidclienteproveedor, d.cfolio ");

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);
            x.mReporteComisiones();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime lfecha = dateTimePicker6.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');


            DateTime lfecha2 = dateTimePicker5.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');

            // string lquery;



            StringBuilder lquery = new StringBuilder();

            lquery.Append("select  c.CCODIGOCLIENTE, c.CRAZONSOCIAL, c.CDIASCREDITOCLIENTE, a.CNOMBREAGENTE, d.CSERIEDOCUMENTO, d.CFOLIO ");
            lquery.Append(",d.cfecha as cfechad, p.cfecha as cfechap, d.CTOTAL, d.CPENDIENTE , isnull(datediff(day,d.cfecha, p.cfecha),-1) as numdias ");
            lquery.Append(", m.CUNIDADESCAPTURADAS, u.CABREVIATURA, pr.CCODIGOPRODUCTO, pr.CNOMBREPRODUCTO, m.ctotal as ctotalmov ");
            lquery.Append(",a.CTIPOAGENTE ");
            lquery.Append(", case  ");
            lquery.Append("when clag.CCODIGOVALORCLASIFICACION  = '1' then isnull(pr.CIMPORTEEXTRA2,0) ");
            lquery.Append("when clag.CCODIGOVALORCLASIFICACION  = '2' then isnull(pr.CIMPORTEEXTRA3,0) ");
            lquery.Append("when clag.CCODIGOVALORCLASIFICACION  = '3' then isnull(pr.CIMPORTEEXTRA4,0) ");
            lquery.Append("else 0 ");
            lquery.Append("end as comision  ");
            lquery.Append(", cl.CVALORCLASIFICACION ");
            lquery.Append(" from admdocumentos d ");
            lquery.Append("join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
            lquery.Append("join admAgentes a on a.CIDAGENTE = d.CIDAGENTE ");
            lquery.Append("left join admAsocCargosAbonos ca on ca.CIDDOCUMENTOCARGO = d.CIDDOCUMENTO ");
            lquery.Append("left join admDocumentos p on ca.CIDDOCUMENTOABONO = p.CIDDOCUMENTO ");
            lquery.Append("join admMovimientos m on m.CIDDOCUMENTO = d.CIDDOCUMENTO ");
            lquery.Append("join admProductos pr on pr.CIDPRODUCTO = m.CIDPRODUCTO ");
            lquery.Append("join admUnidadesMedidaPeso u on m.CIDUNIDAD = u.CIDUNIDAD ");
            lquery.Append("join admClasificacionesValores cl on cl.CIDVALORCLASIFICACION = pr.CIDVALORCLASIFICACION1 ");
            lquery.Append("join admClasificacionesValores clag on clag.CIDVALORCLASIFICACION = a.CIDVALORCLASIFICACION1 ");
            lquery.Append("where d.CIDDOCUMENTODE = 3 and m.cmovtooculto = 0 ");
            lquery.Append("and d.ccancelado = 0 ");


            lquery.Append("and d.CFECHA >= '" + sfecha1 + "'");
            lquery.Append("and d.CFECHA <= '" + sfecha2 + "'");

        /*    if (codigocatalogocomercial1.mGetCodigo() != "" && codigocatalogocomercial2.mGetCodigo() != "")
            {
                lquery.Append("and c.ccodigocliente >= '" + codigocatalogocomercial1.mGetCodigo() + "'");
                lquery.Append("and c.ccodigocliente <= '" + codigocatalogocomercial2.mGetCodigo() + "'");

            }*/

            if (codigocatalogocomercial5.mGetCodigo() != "" && codigocatalogocomercial6.mGetCodigo() != "")
            {
                lquery.Append("and a.ccodigoagente >= '" + codigocatalogocomercial5.mGetCodigo() + "'");
                lquery.Append("and a.ccodigoagente <= '" + codigocatalogocomercial6.mGetCodigo() + "'");

            }

            if (textBox1.Text != "")
            {
                lquery.Append("and d.cseriedocumento ='" + textBox1.Text + "'");
                //lquery.Append("order by Fecha ");
                
            }
            lquery.Append("order by d.cidclienteproveedor, d.cfolio ");

            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);
            x.mReporteReporteComisiones();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            
            DateTime lfecha = dateTimePicker7.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');


            
            // string lquery;



            StringBuilder lquery = new StringBuilder();

         /*   lquery.Append("select d.CIDDOCUMENTODE, CSERIEDOCUMENTO, CFOLIO, a.CCODIGOAGENTE, a.CNOMBREAGENTE, d.CTOTAL, ");
            lquery.Append("case when d.CIDPROYECTO > 0 then p.CNOMBREPROYECTO else case when d.CCANCELADO = 1 then 'Cancelada' else 'Efectivo' END END  " );
            lquery.Append("from admdocumentos d  " );
            lquery.Append("join admAgentes a on d.CIDAGENTE = a.CIDAGENTE  " );
            lquery.Append("join admProyectos p on p.CIDPROYECTO = d.CIDPROYECTO  ");
            lquery.Append("where d.CIDDOCUMENTODE in (3,4)  ") ;*/



            lquery.Append("with mycte (CIDDOCUMENTODE, CSERIEDOCUMENTO, CFOLIO, CCODIGOCLIENTE, CRAZONSOCIAL, dTOTAL, proyecto)  ");
            lquery.Append("as  ");
            lquery.Append("(  ");
            lquery.Append("select d.CIDDOCUMENTODE, CSERIEDOCUMENTO, CFOLIO, c.CCODIGOCLIENTE, c.CRAZONSOCIAL, d.CTOTAL,   ");
	//lquery.Append("case when d.CIDPROYECTO > 0 then p.CNOMBREPROYECTO else case when d.CCANCELADO = 1 then 'Cancelada' else 'Efectivo' END END as proyecto  ");
    lquery.Append("case when d.ccancelado =1 then 'Cancelada' else case when d.CIDPROYECTO > 0 then p.CNOMBREPROYECTO else 'Efectivo' END END as proyecto    ");
	lquery.Append("from admdocumentos d  ");
	//lquery.Append("join admAgentes a on d.CIDAGENTE = a.CIDAGENTE  ");
    lquery.Append("join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
	lquery.Append("join admProyectos p on p.CIDPROYECTO = d.CIDPROYECTO  ");
	lquery.Append("where d.CIDDOCUMENTODE in (3,4)  ");

    if (codigocatalogocomercial7.mGetCodigo() != "")
    {
        //lquery.Append("and d.cusuario = '" + codigocatalogocomercial7.mGetCodigo() + "'");
        lquery.Append("AND d.CSERIEDOCUMENTO like '%" + codigocatalogocomercial7.mGetCodigo() + "'");

    }
    lquery.Append("and d.CFECHA = '" + sfecha1 + "'");
    lquery.Append(")");

    lquery.Append("select CIDDOCUMENTODE, CSERIEDOCUMENTO, CFOLIO, CCODIGOCLIENTE, CRAZONSOCIAL, proyecto ,sum(dTOTAL) as total  ");
    lquery.Append("from mycte  ");
    

            
            
            
            

            /*    if (codigocatalogocomercial1.mGetCodigo() != "" && codigocatalogocomercial2.mGetCodigo() != "")
                {
                    lquery.Append("and c.ccodigocliente >= '" + codigocatalogocomercial1.mGetCodigo() + "'");
                    lquery.Append("and c.ccodigocliente <= '" + codigocatalogocomercial2.mGetCodigo() + "'");

                }*/


            lquery.Append("group by grouping sets   ");
            lquery.Append("(  ");
            lquery.Append("(CIDDOCUMENTODE, CSERIEDOCUMENTO, CFOLIO, CCODIGOCLIENTE, CRAZONSOCIAL,  proyecto),  ");
            lquery.Append("(CIDDOCUMENTODE,proyecto)  ");
            lquery.Append(",(proyecto)  ");
            lquery.Append(")  ");
            //lquery.Append("ORDER BY GROUPING(CIDDOCUMENTODE)  ");
            lquery.Append("ORDER BY 1 ; ");

            lquery.Append("select isnull(m.CIDDOCUMENTODE,0), sum(m.CUNIDADES), u.CABREVIATURA ");
            lquery.Append("from admmovimientos m ");
            lquery.Append("join admDocumentos d on m.CIDDOCUMENTO = d.CIDDOCUMENTO ");
            lquery.Append("join admUnidadesMedidaPeso u on u.CIDUNIDAD = m.CIDUNIDAD ");
            lquery.Append("where d.CIDDOCUMENTODE in (3,4)  AND d.CSERIEDOCUMENTO like '%VA'and d.CFECHA = '" + sfecha1 + "'");
            lquery.Append("group by grouping sets ");
            lquery.Append("( ");
            lquery.Append("(m.CIDDOCUMENTODE, u.CABREVIATURA) ");
            lquery.Append(",(u.CABREVIATURA ) ");
            lquery.Append(") ");
            lquery.Append("ORDER by 1 desc ");

          //  lquery.Append("order by d.CIDDOCUMENTODE ") ;
            x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);
            string sFecha = dateTimePicker7.Value.ToString();
            string lmes = dateTimePicker7.Value.Month.ToString().PadLeft(2,'0'); 
            switch  (lmes)
            {
                case "01":
                        lmes = " Enero "; break;
                case "02":
                        lmes = " Febrero "; break;
                case "03":
                        lmes = " Marzo "; break;
                case "04":
                        lmes = " Abril "; break;
                case "05":
                        lmes = " Mayo"; break;
                case "06":
                        lmes = " Junio"; break;
                case "07":
                        lmes = " Julio "; break;
                case "08":
                        lmes = " Agosto "; break;
                case "09":
                        lmes = " Septiembre "; break;
                case "10":
                        lmes = " Octubre "; break;
                case "11":
                        lmes = " Noviembre "; break;
            }

            sFecha = dateTimePicker7.Value.Day + lmes + " del " + dateTimePicker7.Value.Year;

            x.mReporteCorteCaja(codigocatalogocomercial7.mGetNombre(),sFecha);
        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (textBox2.Text != "")
            {
                DateTime lfecha = dateTimePicker7.Value;
                string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

                StringBuilder lquery = new StringBuilder();

                lquery.Append("select c.CCODIGOCLIENTE, c.CRAZONSOCIAL, c.CDIASCREDITOCLIENTE, d.CSERIEDOCUMENTO, d.CFOLIO,d.cfecha, datediff(day,d.cfecha, getdate()) as diasdif,  ");
                lquery.Append("a.CNOMBREAGENTE, b.CNOMBREAGENTE as CUSUARIO, d.ctotal, d.CPENDIENTE, inicial.cnombreagentesaldo ");
                lquery.Append(", isnull(inicial.saldo,0) as saldoinicial ");
                //            lquery.Append(", isnull(lag(cpendiente) over (order by c.cidclienteproveedor) ,0) ");
                //lquery.Append(", isnull(sum(cpendiente) over (partition by c.cidclienteproveedor order by c.cidclienteproveedor),0) as acumulado ");
                lquery.Append(", 0 as acumulado ");
                lquery.Append("from admDocumentos d ");
                lquery.Append("join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
                lquery.Append("join admAgentes a on d.CIDAGENTE = a.CIDAGENTE ");
                lquery.Append("join admAgentes b on b.CIDAGENTE = c.CIDAGENTEVENTA ");
                lquery.Append("left join  ");
                lquery.Append("( ");
                lquery.Append("select c.CIDCLIENTEPROVEEDOR,d.cpendiente as saldo,isnull(a.cnombreagente,'') as cnombreagentesaldo  ");
                lquery.Append("from admDocumentos d ");
                lquery.Append("join admClientes c on d.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
                lquery.Append("left join admAgentes a on d.CREFERENCIA = a.ccodigoagente ");
                lquery.Append("where d.cfecha = '" + textBox2.Text + "0101'");
                lquery.Append("and d.CCANCELADO = 0 ");
                lquery.Append("and d.cpendiente > 0 ");
                lquery.Append("and d.CIDCONCEPTODOCUMENTO = 39 ");
                lquery.Append(") as inicial on inicial.CIDCLIENTEPROVEEDOR = c.CIDCLIENTEPROVEEDOR ");
                lquery.Append("where year(d.cfecha) = " + textBox2.Text);
                lquery.Append(" and d.CCANCELADO = 0 ");
                lquery.Append("and d.cnaturaleza = 0  ");
                lquery.Append("and d.CIDCONCEPTODOCUMENTO <> 39 ");
                lquery.Append("and d.CPENDIENTE > 0 ");

                //  lquery.Append("order by d.CIDDOCUMENTODE ") ;
                x.mTraerInformacionComercial(lquery, empresasComercial1.aliasbdd);

                string sFecha = dateTimePicker7.Value.ToString();
                /*            string lmes = dateTimePicker7.Value.Month.ToString().PadLeft(2, '0');
                            switch (lmes)
                            {
                                case "01":
                                    lmes = " Enero "; break;
                                case "02":
                                    lmes = " Febrero "; break;
                                case "03":
                                    lmes = " Marzo "; break;
                                case "04":
                                    lmes = " Abril "; break;
                                case "05":
                                    lmes = " Mayo"; break;
                                case "06":
                                    lmes = " Junio"; break;
                                case "07":
                                    lmes = " Julio "; break;
                                case "08":
                                    lmes = " Agosto "; break;
                                case "09":
                                    lmes = " Septiembre "; break;
                                case "10":
                                    lmes = " Octubre "; break;
                                case "11":
                                    lmes = " Noviembre "; break;
                            }

                            sFecha = dateTimePicker7.Value.Day + lmes + " del " + dateTimePicker7.Value.Year;*/

                x.mReporteCobranza(codigocatalogocomercial7.mGetNombre(), textBox2.Text);
            }
            else
                MessageBox.Show("Capture ejercicio");
        }

        private void codigocatalogocomercial7_Load(object sender, EventArgs e)
        {

        }

        private void codigocatalogocomercial1_Load(object sender, EventArgs e)
        {

        }

        private void codigocatalogocomercial5_Load(object sender, EventArgs e)
        {

        }
    }
}
