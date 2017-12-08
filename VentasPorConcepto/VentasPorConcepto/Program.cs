using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace VentasPorConcepto
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new InventarioCapas());
            //Application.Run(new ContpaqiReporteSaldosMovtos());
            //Application.Run(new Form1());
            //Application.Run(new ReporteNominas());
            //Application.Run(new FacturaPedidoComercial());
            //Application.Run(new ReporteFotos());
            Application.Run(new Usuarios());
        }
    }
}
