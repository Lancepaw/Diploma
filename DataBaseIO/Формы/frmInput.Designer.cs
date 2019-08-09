namespace DataBaseIO
{
    partial class frmInput
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
            this.txtSurname = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtPatronymic = new System.Windows.Forms.TextBox();
            this.txtPassport = new System.Windows.Forms.TextBox();
            this.txtRate = new System.Windows.Forms.TextBox();
            this.lblSurname = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.lblPatronymic = new System.Windows.Forms.Label();
            this.cmbDuty = new System.Windows.Forms.ComboBox();
            this.btnCommand = new System.Windows.Forms.Button();
            this.lblPassport = new System.Windows.Forms.Label();
            this.lblDuty = new System.Windows.Forms.Label();
            this.lblRate = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSurname
            // 
            this.txtSurname.Location = new System.Drawing.Point(12, 25);
            this.txtSurname.Name = "txtSurname";
            this.txtSurname.Size = new System.Drawing.Size(100, 20);
            this.txtSurname.TabIndex = 0;
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(118, 25);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(97, 20);
            this.txtName.TabIndex = 1;
            // 
            // txtPatronymic
            // 
            this.txtPatronymic.Location = new System.Drawing.Point(221, 25);
            this.txtPatronymic.Name = "txtPatronymic";
            this.txtPatronymic.Size = new System.Drawing.Size(100, 20);
            this.txtPatronymic.TabIndex = 2;
            // 
            // txtPassport
            // 
            this.txtPassport.Location = new System.Drawing.Point(327, 25);
            this.txtPassport.Name = "txtPassport";
            this.txtPassport.Size = new System.Drawing.Size(100, 20);
            this.txtPassport.TabIndex = 3;
            // 
            // txtRate
            // 
            this.txtRate.Location = new System.Drawing.Point(539, 24);
            this.txtRate.Name = "txtRate";
            this.txtRate.Size = new System.Drawing.Size(100, 20);
            this.txtRate.TabIndex = 4;
            this.txtRate.TextChanged += new System.EventHandler(this.txtRate_TextChanged);
            // 
            // lblSurname
            // 
            this.lblSurname.AutoSize = true;
            this.lblSurname.Location = new System.Drawing.Point(9, 9);
            this.lblSurname.Name = "lblSurname";
            this.lblSurname.Size = new System.Drawing.Size(56, 13);
            this.lblSurname.TabIndex = 5;
            this.lblSurname.Text = "Фамилия";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(115, 9);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(29, 13);
            this.lblName.TabIndex = 6;
            this.lblName.Text = "Имя";
            // 
            // lblPatronymic
            // 
            this.lblPatronymic.AutoSize = true;
            this.lblPatronymic.Location = new System.Drawing.Point(218, 9);
            this.lblPatronymic.Name = "lblPatronymic";
            this.lblPatronymic.Size = new System.Drawing.Size(54, 13);
            this.lblPatronymic.TabIndex = 7;
            this.lblPatronymic.Text = "Отчество";
            // 
            // cmbDuty
            // 
            this.cmbDuty.FormattingEnabled = true;
            this.cmbDuty.Location = new System.Drawing.Point(433, 24);
            this.cmbDuty.Name = "cmbDuty";
            this.cmbDuty.Size = new System.Drawing.Size(100, 21);
            this.cmbDuty.TabIndex = 8;
            // 
            // btnCommand
            // 
            this.btnCommand.Location = new System.Drawing.Point(539, 50);
            this.btnCommand.Name = "btnCommand";
            this.btnCommand.Size = new System.Drawing.Size(98, 23);
            this.btnCommand.TabIndex = 9;
            this.btnCommand.Text = "Команда";
            this.btnCommand.UseVisualStyleBackColor = true;
            this.btnCommand.Click += new System.EventHandler(this.btnCommand_Click);
            // 
            // lblPassport
            // 
            this.lblPassport.AutoSize = true;
            this.lblPassport.Location = new System.Drawing.Point(324, 9);
            this.lblPassport.Name = "lblPassport";
            this.lblPassport.Size = new System.Drawing.Size(91, 13);
            this.lblPassport.TabIndex = 10;
            this.lblPassport.Text = "Номер паспорта";
            // 
            // lblDuty
            // 
            this.lblDuty.AutoSize = true;
            this.lblDuty.Location = new System.Drawing.Point(430, 9);
            this.lblDuty.Name = "lblDuty";
            this.lblDuty.Size = new System.Drawing.Size(65, 13);
            this.lblDuty.TabIndex = 11;
            this.lblDuty.Text = "Должность";
            // 
            // lblRate
            // 
            this.lblRate.AutoSize = true;
            this.lblRate.Location = new System.Drawing.Point(536, 8);
            this.lblRate.Name = "lblRate";
            this.lblRate.Size = new System.Drawing.Size(43, 13);
            this.lblRate.TabIndex = 12;
            this.lblRate.Text = "Ставка";
            // 
            // frmInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 82);
            this.Controls.Add(this.lblRate);
            this.Controls.Add(this.lblDuty);
            this.Controls.Add(this.lblPassport);
            this.Controls.Add(this.btnCommand);
            this.Controls.Add(this.cmbDuty);
            this.Controls.Add(this.lblPatronymic);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.lblSurname);
            this.Controls.Add(this.txtRate);
            this.Controls.Add(this.txtPassport);
            this.Controls.Add(this.txtPatronymic);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.txtSurname);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmInput";
            this.Text = "Ввод данных";
            this.Load += new System.EventHandler(this.frmInput_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSurname;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtPatronymic;
        private System.Windows.Forms.TextBox txtPassport;
        private System.Windows.Forms.TextBox txtRate;
        private System.Windows.Forms.Label lblSurname;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Label lblPatronymic;
        private System.Windows.Forms.ComboBox cmbDuty;
        private System.Windows.Forms.Button btnCommand;
        private System.Windows.Forms.Label lblPassport;
        private System.Windows.Forms.Label lblDuty;
        private System.Windows.Forms.Label lblRate;
    }
}