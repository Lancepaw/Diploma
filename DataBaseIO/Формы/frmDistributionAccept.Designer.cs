namespace DataBaseIO
{
    partial class frmDistributionAccept
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
            this.btnAccept = new System.Windows.Forms.Button();
            this.cmbInput = new System.Windows.Forms.ComboBox();
            this.txtAuditory = new System.Windows.Forms.TextBox();
            this.lblDistribution = new System.Windows.Forms.Label();
            this.lblAuditory = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnAccept
            // 
            this.btnAccept.Location = new System.Drawing.Point(12, 91);
            this.btnAccept.Name = "btnAccept";
            this.btnAccept.Size = new System.Drawing.Size(100, 23);
            this.btnAccept.TabIndex = 0;
            this.btnAccept.Text = "Принять";
            this.btnAccept.UseVisualStyleBackColor = true;
            this.btnAccept.Click += new System.EventHandler(this.btnAccept_Click);
            // 
            // cmbInput
            // 
            this.cmbInput.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbInput.FormattingEnabled = true;
            this.cmbInput.Location = new System.Drawing.Point(12, 25);
            this.cmbInput.Name = "cmbInput";
            this.cmbInput.Size = new System.Drawing.Size(424, 21);
            this.cmbInput.TabIndex = 1;
            // 
            // txtAuditory
            // 
            this.txtAuditory.Location = new System.Drawing.Point(12, 65);
            this.txtAuditory.Name = "txtAuditory";
            this.txtAuditory.Size = new System.Drawing.Size(100, 20);
            this.txtAuditory.TabIndex = 2;
            // 
            // lblDistribution
            // 
            this.lblDistribution.AutoSize = true;
            this.lblDistribution.Location = new System.Drawing.Point(12, 9);
            this.lblDistribution.Name = "lblDistribution";
            this.lblDistribution.Size = new System.Drawing.Size(135, 13);
            this.lblDistribution.TabIndex = 3;
            this.lblDistribution.Text = "Распределение нагрузки";
            // 
            // lblAuditory
            // 
            this.lblAuditory.AutoSize = true;
            this.lblAuditory.Location = new System.Drawing.Point(12, 49);
            this.lblAuditory.Name = "lblAuditory";
            this.lblAuditory.Size = new System.Drawing.Size(60, 13);
            this.lblAuditory.TabIndex = 4;
            this.lblAuditory.Text = "Аудитория";
            // 
            // frmDistributionAccept
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 121);
            this.Controls.Add(this.lblAuditory);
            this.Controls.Add(this.lblDistribution);
            this.Controls.Add(this.txtAuditory);
            this.Controls.Add(this.cmbInput);
            this.Controls.Add(this.btnAccept);
            this.Name = "frmDistributionAccept";
            this.Text = "Назначение нагрузки в расписание";
            this.Load += new System.EventHandler(this.frmDistributionAccept_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAccept;
        private System.Windows.Forms.ComboBox cmbInput;
        private System.Windows.Forms.TextBox txtAuditory;
        private System.Windows.Forms.Label lblDistribution;
        private System.Windows.Forms.Label lblAuditory;
    }
}