namespace DataBaseIO
{
    partial class frmDopWork
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
            this.cmbSemList = new System.Windows.Forms.ComboBox();
            this.cmbLectList = new System.Windows.Forms.ComboBox();
            this.txtWork = new System.Windows.Forms.TextBox();
            this.optNIR = new System.Windows.Forms.RadioButton();
            this.optUMR = new System.Windows.Forms.RadioButton();
            this.optOMR = new System.Windows.Forms.RadioButton();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtComment = new System.Windows.Forms.TextBox();
            this.lblWork = new System.Windows.Forms.Label();
            this.lblComment = new System.Windows.Forms.Label();
            this.lblVolume = new System.Windows.Forms.Label();
            this.txtVolume = new System.Windows.Forms.TextBox();
            this.lblDate = new System.Windows.Forms.Label();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(595, 300);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // cmbSemList
            // 
            this.cmbSemList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemList.FormattingEnabled = true;
            this.cmbSemList.Location = new System.Drawing.Point(12, 12);
            this.cmbSemList.Name = "cmbSemList";
            this.cmbSemList.Size = new System.Drawing.Size(91, 21);
            this.cmbSemList.TabIndex = 1;
            this.cmbSemList.SelectedIndexChanged += new System.EventHandler(this.cmbSemList_SelectedIndexChanged);
            // 
            // cmbLectList
            // 
            this.cmbLectList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLectList.FormattingEnabled = true;
            this.cmbLectList.Location = new System.Drawing.Point(109, 12);
            this.cmbLectList.Name = "cmbLectList";
            this.cmbLectList.Size = new System.Drawing.Size(121, 21);
            this.cmbLectList.TabIndex = 2;
            this.cmbLectList.SelectedIndexChanged += new System.EventHandler(this.cmbLectList_SelectedIndexChanged);
            // 
            // txtWork
            // 
            this.txtWork.Location = new System.Drawing.Point(12, 58);
            this.txtWork.Multiline = true;
            this.txtWork.Name = "txtWork";
            this.txtWork.Size = new System.Drawing.Size(432, 171);
            this.txtWork.TabIndex = 3;
            this.txtWork.TextChanged += new System.EventHandler(this.txtWork_TextChanged);
            // 
            // optNIR
            // 
            this.optNIR.AutoSize = true;
            this.optNIR.Location = new System.Drawing.Point(12, 300);
            this.optNIR.Name = "optNIR";
            this.optNIR.Size = new System.Drawing.Size(48, 17);
            this.optNIR.TabIndex = 4;
            this.optNIR.TabStop = true;
            this.optNIR.Text = "НИР";
            this.optNIR.UseVisualStyleBackColor = true;
            this.optNIR.CheckedChanged += new System.EventHandler(this.optNIR_CheckedChanged);
            // 
            // optUMR
            // 
            this.optUMR.AutoSize = true;
            this.optUMR.Location = new System.Drawing.Point(66, 300);
            this.optUMR.Name = "optUMR";
            this.optUMR.Size = new System.Drawing.Size(49, 17);
            this.optUMR.TabIndex = 5;
            this.optUMR.TabStop = true;
            this.optUMR.Text = "УМР";
            this.optUMR.UseVisualStyleBackColor = true;
            this.optUMR.CheckedChanged += new System.EventHandler(this.optUMR_CheckedChanged);
            // 
            // optOMR
            // 
            this.optOMR.AutoSize = true;
            this.optOMR.Location = new System.Drawing.Point(121, 300);
            this.optOMR.Name = "optOMR";
            this.optOMR.Size = new System.Drawing.Size(49, 17);
            this.optOMR.TabIndex = 6;
            this.optOMR.TabStop = true;
            this.optOMR.Text = "ОМР";
            this.optOMR.UseVisualStyleBackColor = true;
            this.optOMR.CheckedChanged += new System.EventHandler(this.optOMR_CheckedChanged);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(514, 300);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtComment
            // 
            this.txtComment.Location = new System.Drawing.Point(450, 58);
            this.txtComment.Multiline = true;
            this.txtComment.Name = "txtComment";
            this.txtComment.Size = new System.Drawing.Size(220, 171);
            this.txtComment.TabIndex = 8;
            this.txtComment.TextChanged += new System.EventHandler(this.txtComment_TextChanged);
            // 
            // lblWork
            // 
            this.lblWork.AutoSize = true;
            this.lblWork.Location = new System.Drawing.Point(12, 42);
            this.lblWork.Name = "lblWork";
            this.lblWork.Size = new System.Drawing.Size(115, 13);
            this.lblWork.TabIndex = 9;
            this.lblWork.Text = "Наименование работ";
            // 
            // lblComment
            // 
            this.lblComment.AutoSize = true;
            this.lblComment.Location = new System.Drawing.Point(447, 42);
            this.lblComment.Name = "lblComment";
            this.lblComment.Size = new System.Drawing.Size(125, 13);
            this.lblComment.TabIndex = 10;
            this.lblComment.Text = "Отметки о выполнении";
            // 
            // lblVolume
            // 
            this.lblVolume.AutoSize = true;
            this.lblVolume.Location = new System.Drawing.Point(12, 232);
            this.lblVolume.Name = "lblVolume";
            this.lblVolume.Size = new System.Drawing.Size(42, 13);
            this.lblVolume.TabIndex = 11;
            this.lblVolume.Text = "Объём";
            // 
            // txtVolume
            // 
            this.txtVolume.Location = new System.Drawing.Point(12, 248);
            this.txtVolume.Multiline = true;
            this.txtVolume.Name = "txtVolume";
            this.txtVolume.Size = new System.Drawing.Size(432, 46);
            this.txtVolume.TabIndex = 12;
            this.txtVolume.TextChanged += new System.EventHandler(this.txtVolume_TextChanged);
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Location = new System.Drawing.Point(447, 232);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(38, 13);
            this.lblDate.TabIndex = 13;
            this.lblDate.Text = "Сроки";
            // 
            // txtDate
            // 
            this.txtDate.Location = new System.Drawing.Point(450, 248);
            this.txtDate.Multiline = true;
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(220, 46);
            this.txtDate.TabIndex = 14;
            this.txtDate.TextChanged += new System.EventHandler(this.txtDate_TextChanged);
            // 
            // frmDopWork
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(682, 329);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.lblDate);
            this.Controls.Add(this.txtVolume);
            this.Controls.Add(this.lblVolume);
            this.Controls.Add(this.lblComment);
            this.Controls.Add(this.lblWork);
            this.Controls.Add(this.txtComment);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.optOMR);
            this.Controls.Add(this.optUMR);
            this.Controls.Add(this.optNIR);
            this.Controls.Add(this.txtWork);
            this.Controls.Add(this.cmbLectList);
            this.Controls.Add(this.cmbSemList);
            this.Controls.Add(this.btnClose);
            this.Name = "frmDopWork";
            this.Text = "Редактирование дополнительной работы преподавателей";
            this.Load += new System.EventHandler(this.frmDopWork_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ComboBox cmbSemList;
        private System.Windows.Forms.ComboBox cmbLectList;
        private System.Windows.Forms.TextBox txtWork;
        private System.Windows.Forms.RadioButton optNIR;
        private System.Windows.Forms.RadioButton optUMR;
        private System.Windows.Forms.RadioButton optOMR;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtComment;
        private System.Windows.Forms.Label lblWork;
        private System.Windows.Forms.Label lblComment;
        private System.Windows.Forms.Label lblVolume;
        private System.Windows.Forms.TextBox txtVolume;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.TextBox txtDate;
    }
}