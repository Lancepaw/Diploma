namespace DataBaseIO
{
    partial class frmLectInput
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
            this.lblLecturer = new System.Windows.Forms.Label();
            this.cmbLecturer = new System.Windows.Forms.ComboBox();
            this.btnEnter = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblLecturer
            // 
            this.lblLecturer.AutoSize = true;
            this.lblLecturer.Location = new System.Drawing.Point(12, 9);
            this.lblLecturer.Name = "lblLecturer";
            this.lblLecturer.Size = new System.Drawing.Size(131, 13);
            this.lblLecturer.TabIndex = 0;
            this.lblLecturer.Text = "Выбрать преподавателя";
            // 
            // cmbLecturer
            // 
            this.cmbLecturer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturer.FormattingEnabled = true;
            this.cmbLecturer.Location = new System.Drawing.Point(15, 25);
            this.cmbLecturer.Name = "cmbLecturer";
            this.cmbLecturer.Size = new System.Drawing.Size(257, 21);
            this.cmbLecturer.TabIndex = 1;
            // 
            // btnEnter
            // 
            this.btnEnter.Location = new System.Drawing.Point(197, 52);
            this.btnEnter.Name = "btnEnter";
            this.btnEnter.Size = new System.Drawing.Size(75, 23);
            this.btnEnter.TabIndex = 2;
            this.btnEnter.Text = "Ввод";
            this.btnEnter.UseVisualStyleBackColor = true;
            this.btnEnter.Click += new System.EventHandler(this.btnEnter_Click);
            // 
            // frmLectInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 85);
            this.Controls.Add(this.btnEnter);
            this.Controls.Add(this.cmbLecturer);
            this.Controls.Add(this.lblLecturer);
            this.Name = "frmLectInput";
            this.Text = "Выбор преподавателя";
            this.Load += new System.EventHandler(this.frmLectInput_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblLecturer;
        private System.Windows.Forms.ComboBox cmbLecturer;
        private System.Windows.Forms.Button btnEnter;
    }
}