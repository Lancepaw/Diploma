namespace DataBaseIO
{
    partial class frmEditFaculty
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
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.cmbFacultyList = new System.Windows.Forms.ComboBox();
            this.lblFacultyList = new System.Windows.Forms.Label();
            this.lblFaculty = new System.Windows.Forms.Label();
            this.txtFaculty = new System.Windows.Forms.TextBox();
            this.lblShort = new System.Windows.Forms.Label();
            this.txtShort = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDiff = new System.Windows.Forms.Label();
            this.txtDiff = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(197, 227);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(197, 198);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(116, 198);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(116, 227);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // cmbFacultyList
            // 
            this.cmbFacultyList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFacultyList.FormattingEnabled = true;
            this.cmbFacultyList.Location = new System.Drawing.Point(12, 25);
            this.cmbFacultyList.Name = "cmbFacultyList";
            this.cmbFacultyList.Size = new System.Drawing.Size(260, 21);
            this.cmbFacultyList.TabIndex = 4;
            this.cmbFacultyList.SelectedIndexChanged += new System.EventHandler(this.cmbFacultyList_SelectedIndexChanged);
            // 
            // lblFacultyList
            // 
            this.lblFacultyList.AutoSize = true;
            this.lblFacultyList.Location = new System.Drawing.Point(12, 9);
            this.lblFacultyList.Name = "lblFacultyList";
            this.lblFacultyList.Size = new System.Drawing.Size(112, 13);
            this.lblFacultyList.TabIndex = 5;
            this.lblFacultyList.Text = "Список факультетов";
            // 
            // lblFaculty
            // 
            this.lblFaculty.AutoSize = true;
            this.lblFaculty.Location = new System.Drawing.Point(12, 49);
            this.lblFaculty.Name = "lblFaculty";
            this.lblFaculty.Size = new System.Drawing.Size(63, 13);
            this.lblFaculty.TabIndex = 6;
            this.lblFaculty.Text = "Факультет";
            // 
            // txtFaculty
            // 
            this.txtFaculty.Location = new System.Drawing.Point(12, 65);
            this.txtFaculty.Name = "txtFaculty";
            this.txtFaculty.Size = new System.Drawing.Size(260, 20);
            this.txtFaculty.TabIndex = 7;
            // 
            // lblShort
            // 
            this.lblShort.AutoSize = true;
            this.lblShort.Location = new System.Drawing.Point(12, 88);
            this.lblShort.Name = "lblShort";
            this.lblShort.Size = new System.Drawing.Size(128, 13);
            this.lblShort.TabIndex = 8;
            this.lblShort.Text = "Сокращённое название";
            // 
            // txtShort
            // 
            this.txtShort.Location = new System.Drawing.Point(12, 104);
            this.txtShort.Name = "txtShort";
            this.txtShort.Size = new System.Drawing.Size(260, 20);
            this.txtShort.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 232);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Всё готово";
            // 
            // lblDiff
            // 
            this.lblDiff.AutoSize = true;
            this.lblDiff.Location = new System.Drawing.Point(12, 127);
            this.lblDiff.Name = "lblDiff";
            this.lblDiff.Size = new System.Drawing.Size(144, 13);
            this.lblDiff.TabIndex = 11;
            this.lblDiff.Text = "Образовательная система";
            // 
            // txtDiff
            // 
            this.txtDiff.Location = new System.Drawing.Point(12, 143);
            this.txtDiff.Name = "txtDiff";
            this.txtDiff.Size = new System.Drawing.Size(260, 20);
            this.txtDiff.TabIndex = 12;
            // 
            // frmEditFaculty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.txtDiff);
            this.Controls.Add(this.lblDiff);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtShort);
            this.Controls.Add(this.lblShort);
            this.Controls.Add(this.txtFaculty);
            this.Controls.Add(this.lblFaculty);
            this.Controls.Add(this.lblFacultyList);
            this.Controls.Add(this.cmbFacultyList);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditFaculty";
            this.Text = "Редактор факультетов";
            this.Load += new System.EventHandler(this.frmEditFaculty_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.ComboBox cmbFacultyList;
        private System.Windows.Forms.Label lblFacultyList;
        private System.Windows.Forms.Label lblFaculty;
        private System.Windows.Forms.TextBox txtFaculty;
        private System.Windows.Forms.Label lblShort;
        private System.Windows.Forms.TextBox txtShort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDiff;
        private System.Windows.Forms.TextBox txtDiff;
    }
}