namespace VentasPorConcepto
{
    partial class NosSerieAmco
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.empresasComercial1 = new Controles.EmpresasComercial();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCodigoProducto = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtNombreAlmacen = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.txtNombreProducto = new System.Windows.Forms.TextBox();
            this.txtCodigoAlmacen = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // empresasComercial1
            // 
            this.empresasComercial1.Location = new System.Drawing.Point(11, 23);
            this.empresasComercial1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.empresasComercial1.Name = "empresasComercial1";
            this.empresasComercial1.Size = new System.Drawing.Size(650, 54);
            this.empresasComercial1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(654, 79);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(62, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "procesar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(162, 82);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(486, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.textBox1.Leave += new System.EventHandler(this.textBox1_Leave);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "No Serie";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Codigo Producto";
            // 
            // txtCodigoProducto
            // 
            this.txtCodigoProducto.Enabled = false;
            this.txtCodigoProducto.Location = new System.Drawing.Point(163, 108);
            this.txtCodigoProducto.Name = "txtCodigoProducto";
            this.txtCodigoProducto.Size = new System.Drawing.Size(199, 20);
            this.txtCodigoProducto.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 167);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Codigo Almacen";
            // 
            // txtNombreAlmacen
            // 
            this.txtNombreAlmacen.Enabled = false;
            this.txtNombreAlmacen.Location = new System.Drawing.Point(163, 188);
            this.txtNombreAlmacen.Name = "txtNombreAlmacen";
            this.txtNombreAlmacen.Size = new System.Drawing.Size(485, 20);
            this.txtNombreAlmacen.TabIndex = 7;
            this.txtNombreAlmacen.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 223);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(132, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Fecha Ultima Modificacion";
            // 
            // textBox4
            // 
            this.textBox4.Enabled = false;
            this.textBox4.Location = new System.Drawing.Point(162, 218);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(109, 20);
            this.textBox4.TabIndex = 9;
            // 
            // txtNombreProducto
            // 
            this.txtNombreProducto.Enabled = false;
            this.txtNombreProducto.Location = new System.Drawing.Point(163, 134);
            this.txtNombreProducto.Name = "txtNombreProducto";
            this.txtNombreProducto.Size = new System.Drawing.Size(485, 20);
            this.txtNombreProducto.TabIndex = 11;
            // 
            // txtCodigoAlmacen
            // 
            this.txtCodigoAlmacen.Enabled = false;
            this.txtCodigoAlmacen.Location = new System.Drawing.Point(162, 162);
            this.txtCodigoAlmacen.Name = "txtCodigoAlmacen";
            this.txtCodigoAlmacen.Size = new System.Drawing.Size(200, 20);
            this.txtCodigoAlmacen.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(21, 141);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Nombre Producto";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(23, 195);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(90, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Nombre Producto";
            // 
            // NosSerieAmco
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(725, 259);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtCodigoAlmacen);
            this.Controls.Add(this.txtNombreProducto);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtNombreAlmacen);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCodigoProducto);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.empresasComercial1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NosSerieAmco";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Consulta Numero de Serie ";
            this.Load += new System.EventHandler(this.NosSerieAmco_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Controles.EmpresasComercial empresasComercial1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCodigoProducto;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtNombreAlmacen;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox txtNombreProducto;
        private System.Windows.Forms.TextBox txtCodigoAlmacen;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
    }
}