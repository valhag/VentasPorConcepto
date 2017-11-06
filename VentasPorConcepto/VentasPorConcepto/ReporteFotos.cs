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
    public partial class ReporteFotos : Form
    {
        Class1 x = new Class1();
        public ReporteFotos()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            x.mTestFotos();
        }
    }
}
