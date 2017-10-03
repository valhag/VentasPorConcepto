using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VentasPorConcepto
{
    public partial class FacturaPedidoComercial : Form
    {
        public FacturaPedidoComercial()
        {
            InitializeComponent();
        }

        private void FacturaPedidoComercial_Load(object sender, EventArgs e)
        {
            /*
             * select a.CCODIGOAGENTE, a.CNOMBREAGENTE, dp.cfolio, d.CFECHA,d.CFOLIO,p.CCODIGOPRODUCTO, p.CNOMBREPRODUCTO, 
c.cvalorclasificacion, mon.CNOMBREMONEDA, m.CPRECIOCAPTURADO, m.CUNIDADESCAPTURADAS, m.cneto, m.ctotal, d.CFECHAVENCIMIENTO
from admmovimientos m
join admProductos p  on m.CIDPRODUCTO = p.CIDPRODUCTO
join admDocumentos d on d.CIDDOCUMENTO = m.CIDDOCUMENTO
join admAgentes a on a.CIDAGENTE = d.CIDAGENTE
join admClasificacionesValores c on c.CIDVALORCLASIFICACION = p.CIDVALORCLASIFICACION1
join admMonedas mon on mon.CIDMONEDA = d.CIDMONEDA
left join admMovimientos mp on mp.CIDMOVIMIENTO = m.CIDMOVTOORIGEN
and mp.CIDDOCUMENTODE = 2
left join admDocumentos dp on mp.CIDDOCUMENTO = dp.CIDDOCUMENTO
and c.CIDCLASIFICACION = 25
where m.CIDDOCUMENTODE = 4

select * from admClasificaciones order by CIDCLASIFICACION

select * from admClasificacionesValores

select * from admMonedas
             * */
        }



    }
}
