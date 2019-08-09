namespace DataBaseIO
{
    partial class frmSwapScheduleElement
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
            this.btnApply = new System.Windows.Forms.Button();
            this.lblLecturer = new System.Windows.Forms.Label();
            this.cmbLecturer = new System.Windows.Forms.ComboBox();
            this.lblWeekDay = new System.Windows.Forms.Label();
            this.cmbWeekDay = new System.Windows.Forms.ComboBox();
            this.lblPairTime = new System.Windows.Forms.Label();
            this.cmbPairTime = new System.Windows.Forms.ComboBox();
            this.lblWeek = new System.Windows.Forms.Label();
            this.cmbWeek = new System.Windows.Forms.ComboBox();
            this.lblSemestr = new System.Windows.Forms.Label();
            this.cmbSemestr = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(96, 101);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(15, 101);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 23);
            this.btnApply.TabIndex = 1;
            this.btnApply.Text = "Применить";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // lblLecturer
            // 
            this.lblLecturer.AutoSize = true;
            this.lblLecturer.Location = new System.Drawing.Point(12, 9);
            this.lblLecturer.Name = "lblLecturer";
            this.lblLecturer.Size = new System.Drawing.Size(86, 13);
            this.lblLecturer.TabIndex = 2;
            this.lblLecturer.Text = "Преподаватель";
            // 
            // cmbLecturer
            // 
            this.cmbLecturer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturer.FormattingEnabled = true;
            this.cmbLecturer.Location = new System.Drawing.Point(15, 25);
            this.cmbLecturer.Name = "cmbLecturer";
            this.cmbLecturer.Size = new System.Drawing.Size(121, 21);
            this.cmbLecturer.TabIndex = 3;
            // 
            // lblWeekDay
            // 
            this.lblWeekDay.AutoSize = true;
            this.lblWeekDay.Location = new System.Drawing.Point(142, 9);
            this.lblWeekDay.Name = "lblWeekDay";
            this.lblWeekDay.Size = new System.Drawing.Size(73, 13);
            this.lblWeekDay.TabIndex = 4;
            this.lblWeekDay.Text = "День недели";
            // 
            // cmbWeekDay
            // 
            this.cmbWeekDay.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWeekDay.FormattingEnabled = true;
            this.cmbWeekDay.Location = new System.Drawing.Point(145, 25);
            this.cmbWeekDay.Name = "cmbWeekDay";
            this.cmbWeekDay.Size = new System.Drawing.Size(121, 21);
            this.cmbWeekDay.TabIndex = 5;
            // 
            // lblPairTime
            // 
            this.lblPairTime.AutoSize = true;
            this.lblPairTime.Location = new System.Drawing.Point(269, 9);
            this.lblPairTime.Name = "lblPairTime";
            this.lblPairTime.Size = new System.Drawing.Size(84, 13);
            this.lblPairTime.TabIndex = 6;
            this.lblPairTime.Text = "Время занятия";
            // 
            // cmbPairTime
            // 
            this.cmbPairTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPairTime.FormattingEnabled = true;
            this.cmbPairTime.Location = new System.Drawing.Point(272, 25);
            this.cmbPairTime.Name = "cmbPairTime";
            this.cmbPairTime.Size = new System.Drawing.Size(121, 21);
            this.cmbPairTime.TabIndex = 7;
            // 
            // lblWeek
            // 
            this.lblWeek.AutoSize = true;
            this.lblWeek.Location = new System.Drawing.Point(12, 49);
            this.lblWeek.Name = "lblWeek";
            this.lblWeek.Size = new System.Drawing.Size(45, 13);
            this.lblWeek.TabIndex = 8;
            this.lblWeek.Text = "Неделя";
            // 
            // cmbWeek
            // 
            this.cmbWeek.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWeek.FormattingEnabled = true;
            this.cmbWeek.Location = new System.Drawing.Point(15, 65);
            this.cmbWeek.Name = "cmbWeek";
            this.cmbWeek.Size = new System.Drawing.Size(121, 21);
            this.cmbWeek.TabIndex = 9;
            // 
            // lblSemestr
            // 
            this.lblSemestr.AutoSize = true;
            this.lblSemestr.Location = new System.Drawing.Point(142, 49);
            this.lblSemestr.Name = "lblSemestr";
            this.lblSemestr.Size = new System.Drawing.Size(51, 13);
            this.lblSemestr.TabIndex = 10;
            this.lblSemestr.Text = "Семестр";
            // 
            // cmbSemestr
            // 
            this.cmbSemestr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestr.FormattingEnabled = true;
            this.cmbSemestr.Location = new System.Drawing.Point(145, 65);
            this.cmbSemestr.Name = "cmbSemestr";
            this.cmbSemestr.Size = new System.Drawing.Size(121, 21);
            this.cmbSemestr.TabIndex = 11;
            // 
            // frmSwapScheduleElement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(478, 133);
            this.Controls.Add(this.cmbSemestr);
            this.Controls.Add(this.lblSemestr);
            this.Controls.Add(this.cmbWeek);
            this.Controls.Add(this.lblWeek);
            this.Controls.Add(this.cmbPairTime);
            this.Controls.Add(this.lblPairTime);
            this.Controls.Add(this.cmbWeekDay);
            this.Controls.Add(this.lblWeekDay);
            this.Controls.Add(this.cmbLecturer);
            this.Controls.Add(this.lblLecturer);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnClose);
            this.Name = "frmSwapScheduleElement";
            this.Text = "Обмен элемента расписания";
            this.Load += new System.EventHandler(this.frmSwapScheduleElement_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Label lblLecturer;
        private System.Windows.Forms.ComboBox cmbLecturer;
        private System.Windows.Forms.Label lblWeekDay;
        private System.Windows.Forms.ComboBox cmbWeekDay;
        private System.Windows.Forms.Label lblPairTime;
        private System.Windows.Forms.ComboBox cmbPairTime;
        private System.Windows.Forms.Label lblWeek;
        private System.Windows.Forms.ComboBox cmbWeek;
        private System.Windows.Forms.Label lblSemestr;
        private System.Windows.Forms.ComboBox cmbSemestr;
    }
}