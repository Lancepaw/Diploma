namespace DataBaseIO
{
    partial class frmEditScheduleElement
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
            this.btnSave = new System.Windows.Forms.Button();
            this.lblSubject = new System.Windows.Forms.Label();
            this.cmbSubject = new System.Windows.Forms.ComboBox();
            this.lblSubjectType = new System.Windows.Forms.Label();
            this.cmbSubjectType = new System.Windows.Forms.ComboBox();
            this.lblSpecialisation = new System.Windows.Forms.Label();
            this.cmbSpecialisation = new System.Windows.Forms.ComboBox();
            this.lblKursNum = new System.Windows.Forms.Label();
            this.cmbKursNum = new System.Windows.Forms.ComboBox();
            this.lblAuditory = new System.Windows.Forms.Label();
            this.txtAuditory = new System.Windows.Forms.TextBox();
            this.chkBothWeeks = new System.Windows.Forms.CheckBox();
            this.lblGroup = new System.Windows.Forms.Label();
            this.txtGroup = new System.Windows.Forms.TextBox();
            this.lblStream = new System.Windows.Forms.Label();
            this.txtStream = new System.Windows.Forms.TextBox();
            this.lblCode = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.btnSwap = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(308, 160);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(308, 131);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(12, 9);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(121, 13);
            this.lblSubject.TabIndex = 2;
            this.lblSubject.Text = "Читаемая дисциплина";
            // 
            // cmbSubject
            // 
            this.cmbSubject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubject.FormattingEnabled = true;
            this.cmbSubject.Location = new System.Drawing.Point(15, 25);
            this.cmbSubject.Name = "cmbSubject";
            this.cmbSubject.Size = new System.Drawing.Size(304, 21);
            this.cmbSubject.TabIndex = 3;
            // 
            // lblSubjectType
            // 
            this.lblSubjectType.AutoSize = true;
            this.lblSubjectType.Location = new System.Drawing.Point(12, 49);
            this.lblSubjectType.Name = "lblSubjectType";
            this.lblSubjectType.Size = new System.Drawing.Size(70, 13);
            this.lblSubjectType.TabIndex = 4;
            this.lblSubjectType.Text = "Вид занятия";
            // 
            // cmbSubjectType
            // 
            this.cmbSubjectType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubjectType.FormattingEnabled = true;
            this.cmbSubjectType.Location = new System.Drawing.Point(15, 65);
            this.cmbSubjectType.Name = "cmbSubjectType";
            this.cmbSubjectType.Size = new System.Drawing.Size(118, 21);
            this.cmbSubjectType.TabIndex = 5;
            // 
            // lblSpecialisation
            // 
            this.lblSpecialisation.AutoSize = true;
            this.lblSpecialisation.Location = new System.Drawing.Point(139, 49);
            this.lblSpecialisation.Name = "lblSpecialisation";
            this.lblSpecialisation.Size = new System.Drawing.Size(85, 13);
            this.lblSpecialisation.TabIndex = 6;
            this.lblSpecialisation.Text = "Специальность";
            // 
            // cmbSpecialisation
            // 
            this.cmbSpecialisation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpecialisation.FormattingEnabled = true;
            this.cmbSpecialisation.Location = new System.Drawing.Point(139, 65);
            this.cmbSpecialisation.Name = "cmbSpecialisation";
            this.cmbSpecialisation.Size = new System.Drawing.Size(121, 21);
            this.cmbSpecialisation.TabIndex = 7;
            // 
            // lblKursNum
            // 
            this.lblKursNum.AutoSize = true;
            this.lblKursNum.Location = new System.Drawing.Point(266, 49);
            this.lblKursNum.Name = "lblKursNum";
            this.lblKursNum.Size = new System.Drawing.Size(31, 13);
            this.lblKursNum.TabIndex = 8;
            this.lblKursNum.Text = "Курс";
            // 
            // cmbKursNum
            // 
            this.cmbKursNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbKursNum.FormattingEnabled = true;
            this.cmbKursNum.Location = new System.Drawing.Point(266, 65);
            this.cmbKursNum.Name = "cmbKursNum";
            this.cmbKursNum.Size = new System.Drawing.Size(53, 21);
            this.cmbKursNum.TabIndex = 9;
            // 
            // lblAuditory
            // 
            this.lblAuditory.AutoSize = true;
            this.lblAuditory.Location = new System.Drawing.Point(12, 89);
            this.lblAuditory.Name = "lblAuditory";
            this.lblAuditory.Size = new System.Drawing.Size(60, 13);
            this.lblAuditory.TabIndex = 10;
            this.lblAuditory.Text = "Аудитория";
            // 
            // txtAuditory
            // 
            this.txtAuditory.Location = new System.Drawing.Point(15, 105);
            this.txtAuditory.Name = "txtAuditory";
            this.txtAuditory.Size = new System.Drawing.Size(368, 20);
            this.txtAuditory.TabIndex = 11;
            // 
            // chkBothWeeks
            // 
            this.chkBothWeeks.AutoSize = true;
            this.chkBothWeeks.Location = new System.Drawing.Point(204, 135);
            this.chkBothWeeks.Name = "chkBothWeeks";
            this.chkBothWeeks.Size = new System.Drawing.Size(98, 17);
            this.chkBothWeeks.TabIndex = 12;
            this.chkBothWeeks.Text = "на обе недели";
            this.chkBothWeeks.UseVisualStyleBackColor = true;
            // 
            // lblGroup
            // 
            this.lblGroup.AutoSize = true;
            this.lblGroup.Location = new System.Drawing.Point(325, 49);
            this.lblGroup.Name = "lblGroup";
            this.lblGroup.Size = new System.Drawing.Size(42, 13);
            this.lblGroup.TabIndex = 13;
            this.lblGroup.Text = "Группа";
            // 
            // txtGroup
            // 
            this.txtGroup.Location = new System.Drawing.Point(328, 66);
            this.txtGroup.Name = "txtGroup";
            this.txtGroup.Size = new System.Drawing.Size(55, 20);
            this.txtGroup.TabIndex = 14;
            // 
            // lblStream
            // 
            this.lblStream.AutoSize = true;
            this.lblStream.Location = new System.Drawing.Point(325, 9);
            this.lblStream.Name = "lblStream";
            this.lblStream.Size = new System.Drawing.Size(38, 13);
            this.lblStream.TabIndex = 15;
            this.lblStream.Text = "Поток";
            // 
            // txtStream
            // 
            this.txtStream.Location = new System.Drawing.Point(328, 25);
            this.txtStream.Name = "txtStream";
            this.txtStream.Size = new System.Drawing.Size(55, 20);
            this.txtStream.TabIndex = 16;
            // 
            // lblCode
            // 
            this.lblCode.AutoSize = true;
            this.lblCode.Location = new System.Drawing.Point(12, 135);
            this.lblCode.Name = "lblCode";
            this.lblCode.Size = new System.Drawing.Size(26, 13);
            this.lblCode.TabIndex = 17;
            this.lblCode.Text = "Код";
            // 
            // txtCode
            // 
            this.txtCode.Enabled = false;
            this.txtCode.Location = new System.Drawing.Point(15, 151);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(67, 20);
            this.txtCode.TabIndex = 18;
            // 
            // btnSwap
            // 
            this.btnSwap.Location = new System.Drawing.Point(227, 160);
            this.btnSwap.Name = "btnSwap";
            this.btnSwap.Size = new System.Drawing.Size(75, 23);
            this.btnSwap.TabIndex = 19;
            this.btnSwap.Text = "Перенести";
            this.btnSwap.UseVisualStyleBackColor = true;
            this.btnSwap.Click += new System.EventHandler(this.btnSwap_Click);
            // 
            // frmEditScheduleElement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 190);
            this.Controls.Add(this.btnSwap);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.lblCode);
            this.Controls.Add(this.txtStream);
            this.Controls.Add(this.lblStream);
            this.Controls.Add(this.txtGroup);
            this.Controls.Add(this.lblGroup);
            this.Controls.Add(this.chkBothWeeks);
            this.Controls.Add(this.txtAuditory);
            this.Controls.Add(this.lblAuditory);
            this.Controls.Add(this.cmbKursNum);
            this.Controls.Add(this.lblKursNum);
            this.Controls.Add(this.cmbSpecialisation);
            this.Controls.Add(this.lblSpecialisation);
            this.Controls.Add(this.cmbSubjectType);
            this.Controls.Add(this.lblSubjectType);
            this.Controls.Add(this.cmbSubject);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditScheduleElement";
            this.Text = "Редактирование элемента расписания";
            this.Load += new System.EventHandler(this.frmEditScheduleElement_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.ComboBox cmbSubject;
        private System.Windows.Forms.Label lblSubjectType;
        private System.Windows.Forms.ComboBox cmbSubjectType;
        private System.Windows.Forms.Label lblSpecialisation;
        private System.Windows.Forms.ComboBox cmbSpecialisation;
        private System.Windows.Forms.Label lblKursNum;
        private System.Windows.Forms.ComboBox cmbKursNum;
        private System.Windows.Forms.Label lblAuditory;
        private System.Windows.Forms.TextBox txtAuditory;
        private System.Windows.Forms.CheckBox chkBothWeeks;
        private System.Windows.Forms.Label lblGroup;
        private System.Windows.Forms.TextBox txtGroup;
        private System.Windows.Forms.Label lblStream;
        private System.Windows.Forms.TextBox txtStream;
        private System.Windows.Forms.Label lblCode;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.Button btnSwap;
    }
}