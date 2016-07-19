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
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // btnGenerarGastos
            // 
            this.btnGenerarGastos.Location = new System.Drawing.Point(12, 59);
            this.btnGenerarGastos.Name = "btnGenerarGastos";
            this.btnGenerarGastos.Size = new System.Drawing.Size(253, 45);
            this.btnGenerarGastos.TabIndex = 0;
            this.btnGenerarGastos.Text = "Seleccionar archivos";
            this.btnGenerarGastos.UseVisualStyleBackColor = true;
            this.btnGenerarGastos.Click += new System.EventHandler(this.btnGenerarExcel_Click);
            // 
            // LabelReceptor
            // 
            this.LabelReceptor.AutoSize = true;
            this.LabelReceptor.Location = new System.Drawing.Point(10, 25);
            this.LabelReceptor.Name = "LabelReceptor";
            this.LabelReceptor.Size = new System.Drawing.Size(31, 13);
            this.LabelReceptor.TabIndex = 1;
            this.LabelReceptor.Text = "RFC:";
            // 
            // TextBoxReceptor
            // 
            this.TextBoxReceptor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.TextBoxReceptor.Location = new System.Drawing.Point(44, 22);
            this.TextBoxReceptor.MaxLength = 13;
            this.TextBoxReceptor.Name = "TextBoxReceptor";
            this.TextBoxReceptor.Size = new System.Drawing.Size(127, 20);
            this.TextBoxReceptor.TabIndex = 2;
            // 
            // comboBox
            // 
            this.comboBox.BackColor = System.Drawing.SystemColors.Window;
            this.comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox.Items.AddRange(new object[] {
            "Cliente",
            "Proveedor"});
            this.comboBox.Location = new System.Drawing.Point(181, 22);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(84, 21);
            this.comboBox.TabIndex = 3;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 116);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(275, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(275, 138);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.comboBox);
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
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.StatusStrip statusStrip1;
    }
}

