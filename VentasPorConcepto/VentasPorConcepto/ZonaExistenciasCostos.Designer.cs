namespace VentasPorConcepto
{
    partial class ZonaExistenciasCostos
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
            this.codigocatalogocomercial1 = new Controles.codigocatalogocomercial();
            this.codigocatalogocomercial2 = new Controles.codigocatalogocomercial();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // codigocatalogocomercial1
            // 
            this.codigocatalogocomercial1.Location = new System.Drawing.Point(18, 91);
            this.codigocatalogocomercial1.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial1.Name = "codigocatalogocomercial1";
            this.codigocatalogocomercial1.Size = new System.Drawing.Size(597, 21);
            this.codigocatalogocomercial1.TabIndex = 1;
            // 
            // codigocatalogocomercial2
            // 
            this.codigocatalogocomercial2.Location = new System.Drawing.Point(18, 131);
            this.codigocatalogocomercial2.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial2.Name = "codigocatalogocomercial2";
            this.codigocatalogocomercial2.Size = new System.Drawing.Size(597, 21);
            this.codigocatalogocomercial2.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(18, 219);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(564, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "Consultar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 180);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Fecha Corte";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(107, 174);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(88, 20);
            this.dateTimePicker1.TabIndex = 6;
            // 
            // ZonaExistenciasCostos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 267);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.codigocatalogocomercial2);
            this.Controls.Add(this.codigocatalogocomercial1);
            this.Name = "ZonaExistenciasCostos";
            this.Text = "Form5";
            this.Load += new System.EventHandler(this.Form5_Load);
            this.Controls.SetChildIndex(this.empresasComercial1, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial1, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial2, 0);
            this.Controls.SetChildIndex(this.button1, 0);
            this.Controls.SetChildIndex(this.dateTimePicker1, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Controles.codigocatalogocomercial codigocatalogocomercial1;
        private Controles.codigocatalogocomercial codigocatalogocomercial2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
    }
}