using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VentasPorConcepto
{
    public partial class Catalogo : UserControl
    {
        Class1 x = new Class1();
        public int tipo;

        public void setLabel(string name)
        {
            label5.Text = name;
        }

        public string mRegresarCodigo()
        {
            return textBox1.Text;
        }

        public Catalogo()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Catalogo_Load(object sender, EventArgs e)
        {
            
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                
                ComboBox cb = (ComboBox)this.Parent.Controls.Find("comboBox1", true)[0];
                string regresa = x.mRegresarCatalogoValido(2, textBox1.Text, cb.SelectedValue.ToString());
                if (regresa == "")
                {
                    textBox2.Text = "";
                    MessageBox.Show("Proveedor No Valido");
                }
                else
                    textBox2.Text = regresa.Trim();

            }
        }
        private void mBuscarCatalogo()
        {
 
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
