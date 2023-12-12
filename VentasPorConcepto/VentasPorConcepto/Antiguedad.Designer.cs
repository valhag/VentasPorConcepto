namespace VentasPorConcepto
{
    partial class Antiguedad
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
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.codigocatalogocomercial1 = new Controles.codigocatalogocomercial();
            this.codigocatalogocomercial2 = new Controles.codigocatalogocomercial();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.codigocatalogocomercial3 = new Controles.codigocatalogocomercial();
            this.codigocatalogocomercial4 = new Controles.codigocatalogocomercial();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(24, 208);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(489, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Ejecutar Reporte";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(24, 96);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 21);
            this.label4.TabIndex = 32;
            this.label4.Text = "Fecha Corte";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(120, 97);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(85, 20);
            this.dateTimePicker1.TabIndex = 31;
            this.dateTimePicker1.Value = new System.DateTime(2020, 8, 17, 0, 0, 0, 0);
            // 
            // codigocatalogocomercial1
            // 
            this.codigocatalogocomercial1.Location = new System.Drawing.Point(24, 142);
            this.codigocatalogocomercial1.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial1.Name = "codigocatalogocomercial1";
            this.codigocatalogocomercial1.Size = new System.Drawing.Size(518, 21);
            this.codigocatalogocomercial1.TabIndex = 33;
            // 
            // codigocatalogocomercial2
            // 
            this.codigocatalogocomercial2.Location = new System.Drawing.Point(24, 167);
            this.codigocatalogocomercial2.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial2.Name = "codigocatalogocomercial2";
            this.codigocatalogocomercial2.Size = new System.Drawing.Size(518, 21);
            this.codigocatalogocomercial2.TabIndex = 34;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Location = new System.Drawing.Point(267, 82);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(307, 41);
            this.groupBox1.TabIndex = 35;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Reporte por";
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(219, 13);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(66, 17);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.Text = "Compras";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(103, 14);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(66, 17);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Facturas";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // codigocatalogocomercial3
            // 
            this.codigocatalogocomercial3.Location = new System.Drawing.Point(24, 167);
            this.codigocatalogocomercial3.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial3.Name = "codigocatalogocomercial3";
            this.codigocatalogocomercial3.Size = new System.Drawing.Size(518, 21);
            this.codigocatalogocomercial3.TabIndex = 37;
            // 
            // codigocatalogocomercial4
            // 
            this.codigocatalogocomercial4.Location = new System.Drawing.Point(24, 142);
            this.codigocatalogocomercial4.Margin = new System.Windows.Forms.Padding(2);
            this.codigocatalogocomercial4.Name = "codigocatalogocomercial4";
            this.codigocatalogocomercial4.Size = new System.Drawing.Size(518, 21);
            this.codigocatalogocomercial4.TabIndex = 36;
            // 
            // Antiguedad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(635, 267);
            this.Controls.Add(this.codigocatalogocomercial3);
            this.Controls.Add(this.codigocatalogocomercial4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.codigocatalogocomercial2);
            this.Controls.Add(this.codigocatalogocomercial1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.button1);
            this.Name = "Antiguedad";
            this.Text = "Antiguedad";
            this.Load += new System.EventHandler(this.Antiguedad_Load);
            this.Controls.SetChildIndex(this.empresasComercial1, 0);
            this.Controls.SetChildIndex(this.button1, 0);
            this.Controls.SetChildIndex(this.dateTimePicker1, 0);
            this.Controls.SetChildIndex(this.label4, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial1, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial2, 0);
            this.Controls.SetChildIndex(this.groupBox1, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial4, 0);
            this.Controls.SetChildIndex(this.codigocatalogocomercial3, 0);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private Controles.codigocatalogocomercial codigocatalogocomercial1;
        private Controles.codigocatalogocomercial codigocatalogocomercial2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private Controles.codigocatalogocomercial codigocatalogocomercial3;
        private Controles.codigocatalogocomercial codigocatalogocomercial4;
    }
}