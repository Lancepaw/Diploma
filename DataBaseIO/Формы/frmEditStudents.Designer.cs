namespace DataBaseIO
{
    partial class frmEditStudents
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
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lblStudents = new System.Windows.Forms.Label();
            this.cmbStudentsList = new System.Windows.Forms.ComboBox();
            this.cmbKursList = new System.Windows.Forms.ComboBox();
            this.lblKurs = new System.Windows.Forms.Label();
            this.lblFIO = new System.Windows.Forms.Label();
            this.txtFIO = new System.Windows.Forms.TextBox();
            this.lblTheme = new System.Windows.Forms.Label();
            this.txtTheme = new System.Windows.Forms.TextBox();
            this.lblSpeciality = new System.Windows.Forms.Label();
            this.cmbSpecialityList = new System.Windows.Forms.ComboBox();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.cmbDepartmentList = new System.Windows.Forms.ComboBox();
            this.lblLecturer = new System.Windows.Forms.Label();
            this.cmbLecturerList = new System.Windows.Forms.ComboBox();
            this.chkInPlan = new System.Windows.Forms.CheckBox();
            this.btnCopy = new System.Windows.Forms.Button();
            this.lblRows = new System.Windows.Forms.Label();
            this.txtRows = new System.Windows.Forms.TextBox();
            this.gbFilter = new System.Windows.Forms.GroupBox();
            this.cmbLecturerFilt = new System.Windows.Forms.ComboBox();
            this.chkLecturerFilt = new System.Windows.Forms.CheckBox();
            this.gbMain = new System.Windows.Forms.GroupBox();
            this.chkHoured = new System.Windows.Forms.CheckBox();
            this.gbFilter.SuspendLayout();
            this.gbMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(530, 293);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(449, 293);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 1;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(530, 264);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(449, 264);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lblStudents
            // 
            this.lblStudents.AutoSize = true;
            this.lblStudents.Location = new System.Drawing.Point(12, 9);
            this.lblStudents.Name = "lblStudents";
            this.lblStudents.Size = new System.Drawing.Size(47, 13);
            this.lblStudents.TabIndex = 4;
            this.lblStudents.Text = "Студент";
            // 
            // cmbStudentsList
            // 
            this.cmbStudentsList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbStudentsList.FormattingEnabled = true;
            this.cmbStudentsList.Location = new System.Drawing.Point(12, 25);
            this.cmbStudentsList.Name = "cmbStudentsList";
            this.cmbStudentsList.Size = new System.Drawing.Size(431, 21);
            this.cmbStudentsList.TabIndex = 5;
            this.cmbStudentsList.SelectedIndexChanged += new System.EventHandler(this.cmbStudentsList_SelectedIndexChanged);
            // 
            // cmbKursList
            // 
            this.cmbKursList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbKursList.FormattingEnabled = true;
            this.cmbKursList.Location = new System.Drawing.Point(440, 32);
            this.cmbKursList.Name = "cmbKursList";
            this.cmbKursList.Size = new System.Drawing.Size(125, 21);
            this.cmbKursList.TabIndex = 6;
            // 
            // lblKurs
            // 
            this.lblKurs.AutoSize = true;
            this.lblKurs.Location = new System.Drawing.Point(437, 16);
            this.lblKurs.Name = "lblKurs";
            this.lblKurs.Size = new System.Drawing.Size(31, 13);
            this.lblKurs.TabIndex = 7;
            this.lblKurs.Text = "Курс";
            // 
            // lblFIO
            // 
            this.lblFIO.AutoSize = true;
            this.lblFIO.Location = new System.Drawing.Point(6, 16);
            this.lblFIO.Name = "lblFIO";
            this.lblFIO.Size = new System.Drawing.Size(43, 13);
            this.lblFIO.TabIndex = 8;
            this.lblFIO.Text = "Ф.И.О.";
            // 
            // txtFIO
            // 
            this.txtFIO.Location = new System.Drawing.Point(9, 32);
            this.txtFIO.Name = "txtFIO";
            this.txtFIO.Size = new System.Drawing.Size(422, 20);
            this.txtFIO.TabIndex = 9;
            // 
            // lblTheme
            // 
            this.lblTheme.AutoSize = true;
            this.lblTheme.Location = new System.Drawing.Point(6, 55);
            this.lblTheme.Name = "lblTheme";
            this.lblTheme.Size = new System.Drawing.Size(74, 13);
            this.lblTheme.TabIndex = 10;
            this.lblTheme.Text = "Тема работы";
            // 
            // txtTheme
            // 
            this.txtTheme.Location = new System.Drawing.Point(9, 71);
            this.txtTheme.Name = "txtTheme";
            this.txtTheme.Size = new System.Drawing.Size(422, 20);
            this.txtTheme.TabIndex = 11;
            // 
            // lblSpeciality
            // 
            this.lblSpeciality.AutoSize = true;
            this.lblSpeciality.Location = new System.Drawing.Point(289, 94);
            this.lblSpeciality.Name = "lblSpeciality";
            this.lblSpeciality.Size = new System.Drawing.Size(85, 13);
            this.lblSpeciality.TabIndex = 12;
            this.lblSpeciality.Text = "Специальность";
            // 
            // cmbSpecialityList
            // 
            this.cmbSpecialityList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpecialityList.FormattingEnabled = true;
            this.cmbSpecialityList.Location = new System.Drawing.Point(292, 114);
            this.cmbSpecialityList.Name = "cmbSpecialityList";
            this.cmbSpecialityList.Size = new System.Drawing.Size(139, 21);
            this.cmbSpecialityList.TabIndex = 13;
            // 
            // lblDepartment
            // 
            this.lblDepartment.AutoSize = true;
            this.lblDepartment.Location = new System.Drawing.Point(437, 56);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(52, 13);
            this.lblDepartment.TabIndex = 14;
            this.lblDepartment.Text = "Кафедра";
            // 
            // cmbDepartmentList
            // 
            this.cmbDepartmentList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDepartmentList.FormattingEnabled = true;
            this.cmbDepartmentList.Location = new System.Drawing.Point(440, 72);
            this.cmbDepartmentList.Name = "cmbDepartmentList";
            this.cmbDepartmentList.Size = new System.Drawing.Size(125, 21);
            this.cmbDepartmentList.TabIndex = 15;
            // 
            // lblLecturer
            // 
            this.lblLecturer.AutoSize = true;
            this.lblLecturer.Location = new System.Drawing.Point(6, 94);
            this.lblLecturer.Name = "lblLecturer";
            this.lblLecturer.Size = new System.Drawing.Size(78, 13);
            this.lblLecturer.TabIndex = 16;
            this.lblLecturer.Text = "Руководитель";
            // 
            // cmbLecturerList
            // 
            this.cmbLecturerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturerList.FormattingEnabled = true;
            this.cmbLecturerList.Location = new System.Drawing.Point(9, 114);
            this.cmbLecturerList.Name = "cmbLecturerList";
            this.cmbLecturerList.Size = new System.Drawing.Size(277, 21);
            this.cmbLecturerList.TabIndex = 17;
            // 
            // chkInPlan
            // 
            this.chkInPlan.AutoSize = true;
            this.chkInPlan.Location = new System.Drawing.Point(440, 118);
            this.chkInPlan.Name = "chkInPlan";
            this.chkInPlan.Size = new System.Drawing.Size(102, 17);
            this.chkInPlan.TabIndex = 18;
            this.chkInPlan.Text = "Отобр. в плане";
            this.chkInPlan.UseVisualStyleBackColor = true;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(449, 235);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(75, 23);
            this.btnCopy.TabIndex = 19;
            this.btnCopy.Text = "Копировать";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // lblRows
            // 
            this.lblRows.AutoSize = true;
            this.lblRows.Location = new System.Drawing.Point(449, 9);
            this.lblRows.Name = "lblRows";
            this.lblRows.Size = new System.Drawing.Size(98, 13);
            this.lblRows.TabIndex = 20;
            this.lblRows.Text = "Количество строк";
            // 
            // txtRows
            // 
            this.txtRows.Location = new System.Drawing.Point(452, 26);
            this.txtRows.Name = "txtRows";
            this.txtRows.Size = new System.Drawing.Size(125, 20);
            this.txtRows.TabIndex = 21;
            // 
            // gbFilter
            // 
            this.gbFilter.Controls.Add(this.cmbLecturerFilt);
            this.gbFilter.Controls.Add(this.chkLecturerFilt);
            this.gbFilter.Location = new System.Drawing.Point(12, 207);
            this.gbFilter.Name = "gbFilter";
            this.gbFilter.Size = new System.Drawing.Size(431, 109);
            this.gbFilter.TabIndex = 22;
            this.gbFilter.TabStop = false;
            this.gbFilter.Text = "Фильтрация";
            // 
            // cmbLecturerFilt
            // 
            this.cmbLecturerFilt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturerFilt.FormattingEnabled = true;
            this.cmbLecturerFilt.Location = new System.Drawing.Point(9, 42);
            this.cmbLecturerFilt.Name = "cmbLecturerFilt";
            this.cmbLecturerFilt.Size = new System.Drawing.Size(121, 21);
            this.cmbLecturerFilt.TabIndex = 1;
            this.cmbLecturerFilt.SelectedIndexChanged += new System.EventHandler(this.cmbLecturerFilt_SelectedIndexChanged);
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
            this.chkLecturerFilt.CheckedChanged += new System.EventHandler(this.chkLecturerFilt_CheckedChanged);
            // 
            // gbMain
            // 
            this.gbMain.Controls.Add(this.lblFIO);
            this.gbMain.Controls.Add(this.txtFIO);
            this.gbMain.Controls.Add(this.lblTheme);
            this.gbMain.Controls.Add(this.txtTheme);
            this.gbMain.Controls.Add(this.chkInPlan);
            this.gbMain.Controls.Add(this.cmbLecturerList);
            this.gbMain.Controls.Add(this.lblLecturer);
            this.gbMain.Controls.Add(this.cmbSpecialityList);
            this.gbMain.Controls.Add(this.cmbKursList);
            this.gbMain.Controls.Add(this.cmbDepartmentList);
            this.gbMain.Controls.Add(this.lblKurs);
            this.gbMain.Controls.Add(this.lblDepartment);
            this.gbMain.Controls.Add(this.lblSpeciality);
            this.gbMain.Location = new System.Drawing.Point(12, 52);
            this.gbMain.Name = "gbMain";
            this.gbMain.Size = new System.Drawing.Size(593, 149);
            this.gbMain.TabIndex = 23;
            this.gbMain.TabStop = false;
            this.gbMain.Text = "Основные параметры";
            // 
            // chkHoured
            // 
            this.chkHoured.AutoSize = true;
            this.chkHoured.Location = new System.Drawing.Point(452, 207);
            this.chkHoured.Name = "chkHoured";
            this.chkHoured.Size = new System.Drawing.Size(90, 17);
            this.chkHoured.TabIndex = 24;
            this.chkHoured.Text = "В почасовую";
            this.chkHoured.UseVisualStyleBackColor = true;
            // 
            // frmEditStudents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(617, 328);
            this.Controls.Add(this.chkHoured);
            this.Controls.Add(this.gbMain);
            this.Controls.Add(this.gbFilter);
            this.Controls.Add(this.txtRows);
            this.Controls.Add(this.lblRows);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.cmbStudentsList);
            this.Controls.Add(this.lblStudents);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditStudents";
            this.Text = "Редактирование перечня студентов";
            this.Load += new System.EventHandler(this.frmEditStudents_Load);
            this.gbFilter.ResumeLayout(false);
            this.gbFilter.PerformLayout();
            this.gbMain.ResumeLayout(false);
            this.gbMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label lblStudents;
        private System.Windows.Forms.ComboBox cmbStudentsList;
        private System.Windows.Forms.ComboBox cmbKursList;
        private System.Windows.Forms.Label lblKurs;
        private System.Windows.Forms.Label lblFIO;
        private System.Windows.Forms.TextBox txtFIO;
        private System.Windows.Forms.Label lblTheme;
        private System.Windows.Forms.TextBox txtTheme;
        private System.Windows.Forms.Label lblSpeciality;
        private System.Windows.Forms.ComboBox cmbSpecialityList;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.ComboBox cmbDepartmentList;
        private System.Windows.Forms.Label lblLecturer;
        private System.Windows.Forms.ComboBox cmbLecturerList;
        private System.Windows.Forms.CheckBox chkInPlan;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Label lblRows;
        private System.Windows.Forms.TextBox txtRows;
        private System.Windows.Forms.GroupBox gbFilter;
        private System.Windows.Forms.ComboBox cmbLecturerFilt;
        private System.Windows.Forms.CheckBox chkLecturerFilt;
        private System.Windows.Forms.GroupBox gbMain;
        private System.Windows.Forms.CheckBox chkHoured;
    }
}