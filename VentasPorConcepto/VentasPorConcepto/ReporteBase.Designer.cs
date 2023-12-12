namespace VentasPorConcepto
{
    partial class ReporteBase
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
            this.SuspendLayout();
            // 
            // empresasComercial1
            // 
            this.empresasComercial1.Location = new System.Drawing.Point(11, 23);
            this.empresasComercial1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.empresasComercial1.Name = "empresasComercial1";
            this.empresasComercial1.Size = new System.Drawing.Size(612, 54);
            this.empresasComercial1.TabIndex = 0;
            // 
            // ReporteBase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(651, 261);
            this.Controls.Add(this.empresasComercial1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReporteBase";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ReporteBase";
            this.Load += new System.EventHandler(this.ReporteBase_Load);
            this.ResumeLayout(false);

        }

        #endregion

        
        protected Controles.EmpresasComercial empresasComercial1;
    }
}