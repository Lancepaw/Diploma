namespace DataBaseIO
{
    partial class frmEditSubject
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
            this.btnDel = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblSubjectList = new System.Windows.Forms.Label();
            this.cmbSubjectList = new System.Windows.Forms.ComboBox();
            this.txtSubjectName = new System.Windows.Forms.TextBox();
            this.lblSubjectName = new System.Windows.Forms.Label();
            this.lblSubjectShortName = new System.Windows.Forms.Label();
            this.txtSubjectShortName = new System.Windows.Forms.TextBox();
            this.lblPreferences = new System.Windows.Forms.Label();
            this.txtPreferences = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(634, 152);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(553, 152);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 1;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(553, 123);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(634, 123);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblSubjectList
            // 
            this.lblSubjectList.AutoSize = true;
            this.lblSubjectList.Location = new System.Drawing.Point(12, 9);
            this.lblSubjectList.Name = "lblSubjectList";
            this.lblSubjectList.Size = new System.Drawing.Size(101, 13);
            this.lblSubjectList.TabIndex = 4;
            this.lblSubjectList.Text = "Список дисциплин";
            // 
            // cmbSubjectList
            // 
            this.cmbSubjectList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubjectList.FormattingEnabled = true;
            this.cmbSubjectList.Location = new System.Drawing.Point(12, 25);
            this.cmbSubjectList.Name = "cmbSubjectList";
            this.cmbSubjectList.Size = new System.Drawing.Size(697, 21);
            this.cmbSubjectList.TabIndex = 5;
            this.cmbSubjectList.SelectedIndexChanged += new System.EventHandler(this.cmbSubjectList_SelectedIndexChanged);
            // 
            // txtSubjectName
            // 
            this.txtSubjectName.Location = new System.Drawing.Point(12, 65);
            this.txtSubjectName.Name = "txtSubjectName";
            this.txtSubjectName.Size = new System.Drawing.Size(697, 20);
            this.txtSubjectName.TabIndex = 6;
            // 
            // lblSubjectName
            // 
            this.lblSubjectName.AutoSize = true;
            this.lblSubjectName.Location = new System.Drawing.Point(12, 49);
            this.lblSubjectName.Name = "lblSubjectName";
            this.lblSubjectName.Size = new System.Drawing.Size(161, 13);
            this.lblSubjectName.TabIndex = 7;
            this.lblSubjectName.Text = "Полное название дисциплины";
            // 
            // lblSubjectShortName
            // 
            this.lblSubjectShortName.AutoSize = true;
            this.lblSubjectShortName.Location = new System.Drawing.Point(12, 88);
            this.lblSubjectShortName.Name = "lblSubjectShortName";
            this.lblSubjectShortName.Size = new System.Drawing.Size(193, 13);
            this.lblSubjectShortName.TabIndex = 8;
            this.lblSubjectShortName.Text = "Сокращённое название дисциплины";
            // 
            // txtSubjectShortName
            // 
            this.txtSubjectShortName.Location = new System.Drawing.Point(12, 104);
            this.txtSubjectShortName.Name = "txtSubjectShortName";
            this.txtSubjectShortName.Size = new System.Drawing.Size(193, 20);
            this.txtSubjectShortName.TabIndex = 9;
            // 
            // lblPreferences
            // 
            this.lblPreferences.AutoSize = true;
            this.lblPreferences.Location = new System.Drawing.Point(211, 88);
            this.lblPreferences.Name = "lblPreferences";
            this.lblPreferences.Size = new System.Drawing.Size(157, 13);
            this.lblPreferences.TabIndex = 10;
            this.lblPreferences.Text = "Предпочтения по аудиториям";
            // 
            // txtPreferences
            // 
            this.txtPreferences.Location = new System.Drawing.Point(211, 104);
            this.txtPreferences.Name = "txtPreferences";
            this.txtPreferences.Size = new System.Drawing.Size(272, 20);
            this.txtPreferences.TabIndex = 11;
            // 
            // frmEditSubject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(721, 187);
            this.Controls.Add(this.txtPreferences);
            this.Controls.Add(this.lblPreferences);
            this.Controls.Add(this.txtSubjectShortName);
            this.Controls.Add(this.lblSubjectShortName);
            this.Controls.Add(this.lblSubjectName);
            this.Controls.Add(this.txtSubjectName);
            this.Controls.Add(this.cmbSubjectList);
            this.Controls.Add(this.lblSubjectList);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditSubject";
            this.Text = "Редактор дисциплин";
            this.Load += new System.EventHandler(this.frmEditSubject_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblSubjectList;
        private System.Windows.Forms.ComboBox cmbSubjectList;
        private System.Windows.Forms.TextBox txtSubjectName;
        private System.Windows.Forms.Label lblSubjectName;
        private System.Windows.Forms.Label lblSubjectShortName;
        private System.Windows.Forms.TextBox txtSubjectShortName;
        private System.Windows.Forms.Label lblPreferences;
        private System.Windows.Forms.TextBox txtPreferences;
    }
}