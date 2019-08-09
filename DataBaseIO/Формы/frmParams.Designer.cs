namespace DataBaseIO
{
    partial class frmParams
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
            this.btnClose = new System.Windows.Forms.Button();
            this.lblAverageLoad = new System.Windows.Forms.Label();
            this.txtAverageLoad = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblMinistry = new System.Windows.Forms.Label();
            this.txtParams = new System.Windows.Forms.TextBox();
            this.cmbPrintObjects = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(278, 227);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblAverageLoad
            // 
            this.lblAverageLoad.AutoSize = true;
            this.lblAverageLoad.Location = new System.Drawing.Point(12, 9);
            this.lblAverageLoad.Name = "lblAverageLoad";
            this.lblAverageLoad.Size = new System.Drawing.Size(99, 13);
            this.lblAverageLoad.TabIndex = 1;
            this.lblAverageLoad.Text = "Средняя нагрузка";
            // 
            // txtAverageLoad
            // 
            this.txtAverageLoad.Location = new System.Drawing.Point(12, 25);
            this.txtAverageLoad.Name = "txtAverageLoad";
            this.txtAverageLoad.Size = new System.Drawing.Size(100, 20);
            this.txtAverageLoad.TabIndex = 2;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(278, 198);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "Изменить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblMinistry
            // 
            this.lblMinistry.AutoSize = true;
            this.lblMinistry.Location = new System.Drawing.Point(12, 48);
            this.lblMinistry.Name = "lblMinistry";
            this.lblMinistry.Size = new System.Drawing.Size(83, 13);
            this.lblMinistry.TabIndex = 4;
            this.lblMinistry.Text = "Наименование";
            // 
            // txtParams
            // 
            this.txtParams.Location = new System.Drawing.Point(11, 64);
            this.txtParams.Multiline = true;
            this.txtParams.Name = "txtParams";
            this.txtParams.Size = new System.Drawing.Size(342, 54);
            this.txtParams.TabIndex = 5;
            // 
            // cmbPrintObjects
            // 
            this.cmbPrintObjects.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPrintObjects.FormattingEnabled = true;
            this.cmbPrintObjects.Location = new System.Drawing.Point(232, 24);
            this.cmbPrintObjects.Name = "cmbPrintObjects";
            this.cmbPrintObjects.Size = new System.Drawing.Size(121, 21);
            this.cmbPrintObjects.TabIndex = 6;
            this.cmbPrintObjects.SelectedIndexChanged += new System.EventHandler(this.cmbPrintObjects_SelectedIndexChanged);
            // 
            // frmParams
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 262);
            this.Controls.Add(this.cmbPrintObjects);
            this.Controls.Add(this.txtParams);
            this.Controls.Add(this.lblMinistry);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtAverageLoad);
            this.Controls.Add(this.lblAverageLoad);
            this.Controls.Add(this.btnClose);
            this.Name = "frmParams";
            this.Text = "Параметры распределения";
            this.Load += new System.EventHandler(this.frmParams_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblAverageLoad;
        private System.Windows.Forms.TextBox txtAverageLoad;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblMinistry;
        private System.Windows.Forms.TextBox txtParams;
        private System.Windows.Forms.ComboBox cmbPrintObjects;
    }
}