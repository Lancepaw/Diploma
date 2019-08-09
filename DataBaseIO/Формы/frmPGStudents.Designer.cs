namespace DataBaseIO
{
    partial class frmPGStudents
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
            this.gbMain = new System.Windows.Forms.GroupBox();
            this.gbFilter = new System.Windows.Forms.GroupBox();
            this.cmbPGStudentsList = new System.Windows.Forms.ComboBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.lblPGStudentsList = new System.Windows.Forms.Label();
            this.lblRows = new System.Windows.Forms.Label();
            this.txtRows = new System.Windows.Forms.TextBox();
            this.txtFIO = new System.Windows.Forms.TextBox();
            this.lblFIO = new System.Windows.Forms.Label();
            this.cmbKursList = new System.Windows.Forms.ComboBox();
            this.lblKursList = new System.Windows.Forms.Label();
            this.chkLecturerFilt = new System.Windows.Forms.CheckBox();
            this.cmbLecturerFilt = new System.Windows.Forms.ComboBox();
            this.lblLecturerList = new System.Windows.Forms.Label();
            this.cmbLecturerList = new System.Windows.Forms.ComboBox();
            this.cmbDepartmentList = new System.Windows.Forms.ComboBox();
            this.chkPlan = new System.Windows.Forms.CheckBox();
            this.chkBudget = new System.Windows.Forms.CheckBox();
            this.lblHours = new System.Windows.Forms.Label();
            this.txtHours = new System.Windows.Forms.TextBox();
            this.lblDepartmentList = new System.Windows.Forms.Label();
            this.txtTheme = new System.Windows.Forms.TextBox();
            this.lblTheme = new System.Windows.Forms.Label();
            this.gbMain.SuspendLayout();
            this.gbFilter.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbMain
            // 
            this.gbMain.Controls.Add(this.lblTheme);
            this.gbMain.Controls.Add(this.txtTheme);
            this.gbMain.Controls.Add(this.lblDepartmentList);
            this.gbMain.Controls.Add(this.txtHours);
            this.gbMain.Controls.Add(this.lblHours);
            this.gbMain.Controls.Add(this.chkBudget);
            this.gbMain.Controls.Add(this.chkPlan);
            this.gbMain.Controls.Add(this.cmbDepartmentList);
            this.gbMain.Controls.Add(this.cmbLecturerList);
            this.gbMain.Controls.Add(this.lblLecturerList);
            this.gbMain.Controls.Add(this.lblKursList);
            this.gbMain.Controls.Add(this.cmbKursList);
            this.gbMain.Controls.Add(this.lblFIO);
            this.gbMain.Controls.Add(this.txtFIO);
            this.gbMain.Controls.Add(this.txtRows);
            this.gbMain.Controls.Add(this.lblRows);
            this.gbMain.Controls.Add(this.lblPGStudentsList);
            this.gbMain.Controls.Add(this.cmbPGStudentsList);
            this.gbMain.Location = new System.Drawing.Point(12, 12);
            this.gbMain.Name = "gbMain";
            this.gbMain.Size = new System.Drawing.Size(485, 166);
            this.gbMain.TabIndex = 0;
            this.gbMain.TabStop = false;
            this.gbMain.Text = "Основное";
            // 
            // gbFilter
            // 
            this.gbFilter.Controls.Add(this.cmbLecturerFilt);
            this.gbFilter.Controls.Add(this.chkLecturerFilt);
            this.gbFilter.Location = new System.Drawing.Point(12, 184);
            this.gbFilter.Name = "gbFilter";
            this.gbFilter.Size = new System.Drawing.Size(485, 66);
            this.gbFilter.TabIndex = 1;
            this.gbFilter.TabStop = false;
            this.gbFilter.Text = "Фильтрация";
            // 
            // cmbPGStudentsList
            // 
            this.cmbPGStudentsList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPGStudentsList.FormattingEnabled = true;
            this.cmbPGStudentsList.Location = new System.Drawing.Point(6, 32);
            this.cmbPGStudentsList.Name = "cmbPGStudentsList";
            this.cmbPGStudentsList.Size = new System.Drawing.Size(372, 21);
            this.cmbPGStudentsList.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(422, 256);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(341, 256);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(260, 256);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(75, 23);
            this.btnCopy.TabIndex = 0;
            this.btnCopy.Text = "Копировать";
            this.btnCopy.UseVisualStyleBackColor = true;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(179, 256);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 0;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(98, 256);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 0;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            // 
            // lblPGStudentsList
            // 
            this.lblPGStudentsList.AutoSize = true;
            this.lblPGStudentsList.Location = new System.Drawing.Point(6, 16);
            this.lblPGStudentsList.Name = "lblPGStudentsList";
            this.lblPGStudentsList.Size = new System.Drawing.Size(118, 13);
            this.lblPGStudentsList.TabIndex = 1;
            this.lblPGStudentsList.Text = "Перечень аспирантов";
            // 
            // lblRows
            // 
            this.lblRows.AutoSize = true;
            this.lblRows.Location = new System.Drawing.Point(381, 16);
            this.lblRows.Name = "lblRows";
            this.lblRows.Size = new System.Drawing.Size(98, 13);
            this.lblRows.TabIndex = 2;
            this.lblRows.Text = "Количество строк";
            // 
            // txtRows
            // 
            this.txtRows.Location = new System.Drawing.Point(384, 33);
            this.txtRows.Name = "txtRows";
            this.txtRows.Size = new System.Drawing.Size(95, 20);
            this.txtRows.TabIndex = 3;
            // 
            // txtFIO
            // 
            this.txtFIO.Location = new System.Drawing.Point(6, 72);
            this.txtFIO.Name = "txtFIO";
            this.txtFIO.Size = new System.Drawing.Size(372, 20);
            this.txtFIO.TabIndex = 4;
            // 
            // lblFIO
            // 
            this.lblFIO.AutoSize = true;
            this.lblFIO.Location = new System.Drawing.Point(6, 56);
            this.lblFIO.Name = "lblFIO";
            this.lblFIO.Size = new System.Drawing.Size(43, 13);
            this.lblFIO.TabIndex = 5;
            this.lblFIO.Text = "Ф.И.О.";
            // 
            // cmbKursList
            // 
            this.cmbKursList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbKursList.FormattingEnabled = true;
            this.cmbKursList.Location = new System.Drawing.Point(6, 111);
            this.cmbKursList.Name = "cmbKursList";
            this.cmbKursList.Size = new System.Drawing.Size(43, 21);
            this.cmbKursList.TabIndex = 6;
            // 
            // lblKursList
            // 
            this.lblKursList.AutoSize = true;
            this.lblKursList.Location = new System.Drawing.Point(6, 95);
            this.lblKursList.Name = "lblKursList";
            this.lblKursList.Size = new System.Drawing.Size(31, 13);
            this.lblKursList.TabIndex = 7;
            this.lblKursList.Text = "Курс";
            // 
            // chkLecturerFilt
            // 
            this.chkLecturerFilt.AutoSize = true;
            this.chkLecturerFilt.Location = new System.Drawing.Point(9, 19);
            this.chkLecturerFilt.Name = "chkLecturerFilt";
            this.chkLecturerFilt.Size = new System.Drawing.Size(113, 17);
            this.chkLecturerFilt.TabIndex = 0;
            this.chkLecturerFilt.Text = "по руководителю";
            this.chkLecturerFilt.UseVisualStyleBackColor = true;
            // 
            // cmbLecturerFilt
            // 
            this.cmbLecturerFilt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturerFilt.FormattingEnabled = true;
            this.cmbLecturerFilt.Location = new System.Drawing.Point(9, 39);
            this.cmbLecturerFilt.Name = "cmbLecturerFilt";
            this.cmbLecturerFilt.Size = new System.Drawing.Size(113, 21);
            this.cmbLecturerFilt.TabIndex = 1;
            // 
            // lblLecturerList
            // 
            this.lblLecturerList.AutoSize = true;
            this.lblLecturerList.Location = new System.Drawing.Point(52, 95);
            this.lblLecturerList.Name = "lblLecturerList";
            this.lblLecturerList.Size = new System.Drawing.Size(78, 13);
            this.lblLecturerList.TabIndex = 8;
            this.lblLecturerList.Text = "Руководитель";
            // 
            // cmbLecturerList
            // 
            this.cmbLecturerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturerList.FormattingEnabled = true;
            this.cmbLecturerList.Location = new System.Drawing.Point(55, 111);
            this.cmbLecturerList.Name = "cmbLecturerList";
            this.cmbLecturerList.Size = new System.Drawing.Size(121, 21);
            this.cmbLecturerList.TabIndex = 9;
            // 
            // cmbDepartmentList
            // 
            this.cmbDepartmentList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDepartmentList.FormattingEnabled = true;
            this.cmbDepartmentList.Location = new System.Drawing.Point(240, 138);
            this.cmbDepartmentList.Name = "cmbDepartmentList";
            this.cmbDepartmentList.Size = new System.Drawing.Size(239, 21);
            this.cmbDepartmentList.TabIndex = 10;
            // 
            // chkPlan
            // 
            this.chkPlan.AutoSize = true;
            this.chkPlan.Location = new System.Drawing.Point(384, 74);
            this.chkPlan.Name = "chkPlan";
            this.chkPlan.Size = new System.Drawing.Size(66, 17);
            this.chkPlan.TabIndex = 11;
            this.chkPlan.Text = "В плане";
            this.chkPlan.UseVisualStyleBackColor = true;
            // 
            // chkBudget
            // 
            this.chkBudget.AutoSize = true;
            this.chkBudget.Location = new System.Drawing.Point(6, 138);
            this.chkBudget.Name = "chkBudget";
            this.chkBudget.Size = new System.Drawing.Size(66, 17);
            this.chkBudget.TabIndex = 12;
            this.chkBudget.Text = "Бюджет";
            this.chkBudget.UseVisualStyleBackColor = true;
            // 
            // lblHours
            // 
            this.lblHours.AutoSize = true;
            this.lblHours.Location = new System.Drawing.Point(78, 139);
            this.lblHours.Name = "lblHours";
            this.lblHours.Size = new System.Drawing.Size(35, 13);
            this.lblHours.TabIndex = 13;
            this.lblHours.Text = "Часы";
            // 
            // txtHours
            // 
            this.txtHours.Location = new System.Drawing.Point(119, 136);
            this.txtHours.Name = "txtHours";
            this.txtHours.Size = new System.Drawing.Size(57, 20);
            this.txtHours.TabIndex = 14;
            // 
            // lblDepartmentList
            // 
            this.lblDepartmentList.AutoSize = true;
            this.lblDepartmentList.Location = new System.Drawing.Point(182, 139);
            this.lblDepartmentList.Name = "lblDepartmentList";
            this.lblDepartmentList.Size = new System.Drawing.Size(52, 13);
            this.lblDepartmentList.TabIndex = 15;
            this.lblDepartmentList.Text = "Кафедра";
            // 
            // txtTheme
            // 
            this.txtTheme.Location = new System.Drawing.Point(182, 111);
            this.txtTheme.Name = "txtTheme";
            this.txtTheme.Size = new System.Drawing.Size(297, 20);
            this.txtTheme.TabIndex = 16;
            // 
            // lblTheme
            // 
            this.lblTheme.AutoSize = true;
            this.lblTheme.Location = new System.Drawing.Point(182, 95);
            this.lblTheme.Name = "lblTheme";
            this.lblTheme.Size = new System.Drawing.Size(34, 13);
            this.lblTheme.TabIndex = 17;
            this.lblTheme.Text = "Тема";
            // 
            // frmPGStudents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(509, 292);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.gbFilter);
            this.Controls.Add(this.gbMain);
            this.Name = "frmPGStudents";
            this.Text = "Редактирование перечня аспирантов";
            this.Load += new System.EventHandler(this.frmPGStudents_Load);
            this.gbMain.ResumeLayout(false);
            this.gbMain.PerformLayout();
            this.gbFilter.ResumeLayout(false);
            this.gbFilter.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbMain;
        private System.Windows.Forms.ComboBox cmbPGStudentsList;
        private System.Windows.Forms.GroupBox gbFilter;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.TextBox txtHours;
        private System.Windows.Forms.Label lblHours;
        private System.Windows.Forms.CheckBox chkBudget;
        private System.Windows.Forms.CheckBox chkPlan;
        private System.Windows.Forms.ComboBox cmbDepartmentList;
        private System.Windows.Forms.ComboBox cmbLecturerList;
        private System.Windows.Forms.Label lblLecturerList;
        private System.Windows.Forms.Label lblKursList;
        private System.Windows.Forms.ComboBox cmbKursList;
        private System.Windows.Forms.Label lblFIO;
        private System.Windows.Forms.TextBox txtFIO;
        private System.Windows.Forms.TextBox txtRows;
        private System.Windows.Forms.Label lblRows;
        private System.Windows.Forms.Label lblPGStudentsList;
        private System.Windows.Forms.ComboBox cmbLecturerFilt;
        private System.Windows.Forms.CheckBox chkLecturerFilt;
        private System.Windows.Forms.Label lblDepartmentList;
        private System.Windows.Forms.Label lblTheme;
        private System.Windows.Forms.TextBox txtTheme;
    }
}