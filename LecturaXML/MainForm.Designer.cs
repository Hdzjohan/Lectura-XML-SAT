namespace LecturaXML
{
    partial class MainForm
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
            this.btnGenerarGastos = new System.Windows.Forms.Button();
            this.LabelReceptor = new System.Windows.Forms.Label();
            this.TextBoxReceptor = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnGenerarGastos
            // 
            this.btnGenerarGastos.Location = new System.Drawing.Point(10, 60);
            this.btnGenerarGastos.Name = "btnGenerarGastos";
            this.btnGenerarGastos.Size = new System.Drawing.Size(262, 45);
            this.btnGenerarGastos.TabIndex = 0;
            this.btnGenerarGastos.Text = "Seleccionar archivos";
            this.btnGenerarGastos.UseVisualStyleBackColor = true;
            this.btnGenerarGastos.Click += new System.EventHandler(this.btnGenerarGastos_Click);
            // 
            // LabelReceptor
            // 
            this.LabelReceptor.AutoSize = true;
            this.LabelReceptor.Location = new System.Drawing.Point(10, 25);
            this.LabelReceptor.Name = "LabelReceptor";
            this.LabelReceptor.Size = new System.Drawing.Size(75, 13);
            this.LabelReceptor.TabIndex = 1;
            this.LabelReceptor.Text = "RFC Receptor";
            // 
            // TextBoxReceptor
            // 
            this.TextBoxReceptor.Location = new System.Drawing.Point(91, 23);
            this.TextBoxReceptor.Name = "TextBoxReceptor";
            this.TextBoxReceptor.Size = new System.Drawing.Size(181, 20);
            this.TextBoxReceptor.TabIndex = 2;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 121);
            this.Controls.Add(this.TextBoxReceptor);
            this.Controls.Add(this.LabelReceptor);
            this.Controls.Add(this.btnGenerarGastos);
            this.Name = "MainForm";
            this.Text = "Generación de Gastos";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerarGastos;
        private System.Windows.Forms.Label LabelReceptor;
        private System.Windows.Forms.TextBox TextBoxReceptor;
    }
}

